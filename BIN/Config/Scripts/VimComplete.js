/*===========================================================================
Copyright (c) 2004-2005 Alexander Kuntashov
=============================================================================
Скрипт:  VimComplete.js
Версия:  1.4
Автор:   Александр Кунташов
E-mail:  kuntashov at yandex dot ru
ICQ UIN: 338758861
Описание: 
    Атодополнение слов в стиле редактора Vim
===========================================================================*/
/*
    Макрос NextWord()
            Подбирает часть слова слева от курсора и пытается дополнить его,
        ища вперед по тексту слова, с такой же левой частью. Подставляет 
        первое подходящее. Следующий вызов макроса подставит следующее за 
        первым найденным словом и так далее по кругу (дойдя до последней строки 
        модуля поиск продолжится с первой строки). 
        
    Макрос PrevWord()
            Тоже самое, только поиск слов осуществляется в обратном направлении.
   
    В классическом Vim используются следующие хоткеи:
    
        Ctrl + N  для дополнения с поиском вперед (следующее слово, Next word)
        Ctrl + P  для дополнения с поиском назад  (предыдущее слово, Previous word)
*/

/* ==========================================================================
                                    МАКРОСЫ
========================================================================== */

// следующее соответствие
function NextWord() // Ctrl + N
{
    completeWord(true);
}

// предыдущее соответствие
function PrevWord() // Ctrl + P
{
    completeWord(false);
}

/* ==========================================================================
                                IMPLEMENTATION                       
========================================================================== */

var vat = new VimAutoCompletionTool;

function completeWord(lookForward)
{
    var doc;

    if (!(doc = CommonScripts.GetTextDocIfOpened(false))) {
        return;
    }
    
    vat.setup(doc, lookForward);
    vat.completeWord();
}

/* Возвращает часть слова слева от 
    текущего положения курсора */
function getLeftWord(doc)
{
    var cl, word = '';
    cl = doc.Range(doc.SelStartLine);
    for (var i=doc.SelStartCol-1;
            (i>=0)&&cl.charAt(i).match(/[\wА-Яа-я]/i);
                word = cl.charAt(i--) + word)
        ;    
    return word;
}

/* Примитивный класс для выделения слов в строке, 
их фильтрации и последовательному перебору. 
Написано в лоб, так что больших скоростей не обещаю
(собственно, именно по этому я не разбираю сразу весь 
текст на слова, а делаю это построчно, лишая себя возможности
легко избавиться от повторяющихся соответствий, что мне кажется 
менее критичным в контексте юзабилити)*/

function Line(str)
{
// private 

    var s = str;
    var words = null; 
    var iterator = -1;
    var forward;
    
// public

    this.setOrder = function (lookForward)
    {
        forward = lookForward;
    }
    this.next = function()
    {        
        return this.item(iterator += (forward?1:(-1)));
    }   
    this.prev = function() 
    {
        return this.item(iterator -= (forward?1:(-1)));
    }
    this.word = function()
    {
        return this.item(iterator);
    }    
    this.setup = function (fwd)
    {
        forward = fwd;
        iterator = forward ? -1 : (this.count());
    }
    this.reset = function () 
    {
        words = s.split(/[^\wА-я]+/);
        this.setup(true);
    }
    this.assign = function (ix) 
    {
        return ((typeof(words)=="object")&&(ix>=0)&&(ix<words.length));                  
    }
    this.item = function (ix) 
    {
        if (this.assign(ix)) {
            return words[ix];
        }
    }
    this.count = function ()
    {
        return words.length;        
    }
    this.filter = function (pattern, unique)
    {
        var used = "";
        if (this.assign(0)) {
            var nw = new Array();
            for (var i=0; i<this.count(); i++) {
                if (this.item(i).match(pattern)) {
                    if (unique) {
                        if (!used.match(new RegExp(this.item(i) + ";","i"))) {
                            used += this.item(i) + ";";
                            nw[nw.length] = this.item(i);
                        }
                    }
                    else {
                       nw[nw.length] = this.item(i);
                    }
                }
            }
            words = nw;
            return true;
        }
        return false;
    }        
   
    this.reset(); // инициализация
}

/* По [моим] ощущениям практически повторяет поведение соответсвующих
функций Vim (я про ^P/^N), за исключением относительно
небольшого, но все же заметного минуса - повторяющиеся в 
разных строках соответствия не исключаются из списка 
подстановки. Пока меня это устраивает, хотя бы потому,
что скрипт используется совместно с Телепатом, а не сам по себе. 
Еще один небольшой баг заключается в том, что при дополнении
пустого слова в строку статуса выводится не совсем корректная
информация и к исходному (пустому) слову, к сожалению, не 
вернуться никак, но это совсем уж не существенно, хотя возможно,
как нибудь исправлю для красоты :-) */

function VimAutoCompletionTool(_) // orgy citadel :-)
{
// private

    var srcDoc;    // исходный документ
    var srcLine;   // исходная строка документа
    var srcCol;    // первая позиция в строке перед исходным словом
    var srcWord;   // исходное слово (которое пытаемся дополнить)
    var lastWord;  // последнее использованное в подстановке слово
    var curLineIx; // индекс текущей строки (из которой берутся соответствия)
    var words;     // список слов-соответствий текущей строки
    
    var forwardSearch; // задает направление поиска по тексту: true - вперед, false - назад
    var pattern;  // шаблон (регулярное выражение), описывающий соответствие исходному слову 
    
    var counter; // счетчик соответствий
    var total; // общее число соответствий
        
// public

    /* выполняет (ре)инициализацию объекта, 
        если это необходимо */
    this.setup = function (doc, lookForward)
    {
        var word = getLeftWord(doc);            
        if (this.isNewLoop(doc, word)) { // реинициализация
            srcDoc     = doc;            
            srcLine    = doc.SelStartLine;
            srcCol     = doc.SelStartCol - word.length;            
            srcWord    = word;            
            lastWord   = word; // чтобы корректно сделать первую подстановку
            curLineIx  = srcLine;            
            pattern     = new RegExp("^" + word, "i");     
            // начинаем искать соответствия начиная с исходной строки
            words = this.parseLine(lookForward ? this.rightPart() : this.leftPart());  
            
            counter = 0;
            total = null;
        }               
        forwardSearch = lookForward;
    }
    /* условие необходимости произвести переинициализацию 
        переменных членов объекта VimAutoCompletionTool */
    this.isNewLoop = function (doc, word)
    {
        return !((srcDoc)
            &&(srcDoc.Path == doc.Path) 
            &&(srcLine == doc.SelStartLine) 
            &&(lastWord == word));         
    }
    /* проверяет, не выходит ли индекс строки за допустимые 
        границы */
    this.assign = function (lIx)
    {
        return (srcDoc&&(0<=lIx)&&(lIx<srcDoc.LineCount));
    }
    /* берет следующее соответствие и подставляет его на
        место исходного слова */
    this.completeWord = function ()
    {
        var word;
        while (true) {  
            words.setOrder(forwardSearch);
            word = words.next();
            if (word) {
                this.complete(word);
                return;
            }
            words = this.nextLine();                                                   
        }
    }
    /* строит и возвращает список соответсвующих слов для следующей 
        по порядку строки */
    this.nextLine = function ()
    {                   
        curLineIx += (forwardSearch ? 1 : -1); 
        if (forwardSearch) {
            if (curLineIx == srcDoc.LineCount) {
                curLineIx = 0;
            }
        } 
        else {
            if (curLineIx < 0) {
                curLineIx = srcDoc.LineCount-1;
            }
        }
        return this.parseLine(this.curLine());               
    }
    /* "разбирает" переданную в качестве параметра строку на слова
        и фильтрует их в соотвествии с шаблоном, который описывает
        подходящие соответсвия для исходного слова*/
    this.parseLine = function (src)
    {
        var w = new Line(src);
        w.filter(pattern, true);
        w.setup(forwardSearch);
        return w;
    }
    /* выполняет подстановку очередного соответствия
        вместо исходного слова */
    this.complete = function (word)
    {        
        srcDoc.Range(srcLine) = this.leftPart() + word + this.rightPart();
        srcDoc.MoveCaret(srcLine, srcCol+word.length);
            
        lastWord = word;

        counter += forwardSearch ? 1 : -1;          
        if ((curLineIx == srcLine)&&(lastWord == srcWord)) {
            if ((!total)&&counter) {                                
                total = Math.abs(counter) - 1;                     
            }
            counter = 0;
        }
        // практически совсем как в Vim :-)       
        Status(counter 
            ? "Cоответствие: "+Math.abs(counter).toString()+(total?" из " + total:"")
            : (total ? "Исходное слово" : "Шаблон не найден")); 
    } 
    // возвращает текущую строку
    this.curLine = function ()
    {
        var src = "";
        if (this.assign(curLineIx)) {
            /* поскольку исходная строка у нас постоянно меняется,
                ее "собираем" отдельно, возвращая на место исходное слово */
            if (curLineIx == srcLine) {
                src = this.leftPart() + srcWord + this.rightPart();           
            }
            else {
                src = srcDoc.Range(curLineIx);
            }           
        }
        return src;
    }
    /* левая половина строки, содержащей исходное слово; 
        само исходное слово не включается */
    this.leftPart = function ()
    {
        return srcDoc.Range(srcLine, 0, srcLine, srcCol);
    }
    /* правая половина строки, содержащей исходное слово;
        само исходное слово не включается */
    this.rightPart = function ()
    {      
        return srcDoc.Range(
            srcLine, srcCol+lastWord.length,
            srcLine, srcDoc.LineLen(srcLine) - 1);
    }
}

/*
 * Процедура инициализации скрипта 
 */
function Init(_) // Фиктивный параметр, чтобы процедура не попадала в макросы
{    
    try {
        var c = null;
        if (!(c = new ActiveXObject("OpenConf.CommonServices"))) {
            throw(true); // вызываем исключение
        }        
        c.SetConfig(Configurator);
        SelfScript.AddNamedItem("CommonScripts", c, false);
    }
    catch (e) {
        Message("Не могу создать объект OpenConf.CommonServices", mRedErr);    
        Message("Скрипт " & SelfScript.Name & " не загружен", mInformation);
        Scripts.UnLoad(SelfScript.Name);
    }
}

Init(); // При загрузке скрипта выполняем его инициализацию
