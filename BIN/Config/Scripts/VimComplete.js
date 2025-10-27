/*===========================================================================
Copyright (c) 2004-2005 Alexander Kuntashov
=============================================================================
������:  VimComplete.js
������:  1.4
�����:   ��������� ��������
E-mail:  kuntashov at yandex dot ru
ICQ UIN: 338758861
��������: 
    ������������� ���� � ����� ��������� Vim
===========================================================================*/
/*
    ������ NextWord()
            ��������� ����� ����� ����� �� ������� � �������� ��������� ���,
        ��� ������ �� ������ �����, � ����� �� ����� ������. ����������� 
        ������ ����������. ��������� ����� ������� ��������� ��������� �� 
        ������ ��������� ������ � ��� ����� �� ����� (����� �� ��������� ������ 
        ������ ����� ����������� � ������ ������). 
        
    ������ PrevWord()
            ���� �����, ������ ����� ���� �������������� � �������� �����������.
   
    � ������������ Vim ������������ ��������� ������:
    
        Ctrl + N  ��� ���������� � ������� ������ (��������� �����, Next word)
        Ctrl + P  ��� ���������� � ������� �����  (���������� �����, Previous word)
*/

/* ==========================================================================
                                    �������
========================================================================== */

// ��������� ������������
function NextWord() // Ctrl + N
{
    completeWord(true);
}

// ���������� ������������
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

/* ���������� ����� ����� ����� �� 
    �������� ��������� ������� */
function getLeftWord(doc)
{
    var cl, word = '';
    cl = doc.Range(doc.SelStartLine);
    for (var i=doc.SelStartCol-1;
            (i>=0)&&cl.charAt(i).match(/[\w�-��-�]/i);
                word = cl.charAt(i--) + word)
        ;    
    return word;
}

/* ����������� ����� ��� ��������� ���� � ������, 
�� ���������� � ����������������� ��������. 
�������� � ���, ��� ��� ������� ��������� �� ������
(����������, ������ �� ����� � �� �������� ����� ���� 
����� �� �����, � ����� ��� ���������, ����� ���� �����������
����� ���������� �� ������������� ������������, ��� ��� ������� 
����� ��������� � ��������� ���������)*/

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
        words = s.split(/[^\w�-�]+/);
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
   
    this.reset(); // �������������
}

/* �� [����] ��������� ����������� ��������� ��������� ��������������
������� Vim (� ��� ^P/^N), �� ����������� ������������
����������, �� ��� �� ��������� ������ - ������������� � 
������ ������� ������������ �� ����������� �� ������ 
�����������. ���� ���� ��� ����������, ���� �� ������,
��� ������ ������������ ��������� � ���������, � �� ��� �� ����. 
��� ���� ��������� ��� ����������� � ���, ��� ��� ����������
������� ����� � ������ ������� ��������� �� ������ ����������
���������� � � ��������� (�������) �����, � ���������, �� 
��������� �����, �� ��� ������ �� �� �����������, ���� ��������,
��� ������ �������� ��� ������� :-) */

function VimAutoCompletionTool(_) // orgy citadel :-)
{
// private

    var srcDoc;    // �������� ��������
    var srcLine;   // �������� ������ ���������
    var srcCol;    // ������ ������� � ������ ����� �������� ������
    var srcWord;   // �������� ����� (������� �������� ���������)
    var lastWord;  // ��������� �������������� � ����������� �����
    var curLineIx; // ������ ������� ������ (�� ������� ������� ������������)
    var words;     // ������ ����-������������ ������� ������
    
    var forwardSearch; // ������ ����������� ������ �� ������: true - ������, false - �����
    var pattern;  // ������ (���������� ���������), ����������� ������������ ��������� ����� 
    
    var counter; // ������� ������������
    var total; // ����� ����� ������������
        
// public

    /* ��������� (��)������������� �������, 
        ���� ��� ���������� */
    this.setup = function (doc, lookForward)
    {
        var word = getLeftWord(doc);            
        if (this.isNewLoop(doc, word)) { // ���������������
            srcDoc     = doc;            
            srcLine    = doc.SelStartLine;
            srcCol     = doc.SelStartCol - word.length;            
            srcWord    = word;            
            lastWord   = word; // ����� ��������� ������� ������ �����������
            curLineIx  = srcLine;            
            pattern     = new RegExp("^" + word, "i");     
            // �������� ������ ������������ ������� � �������� ������
            words = this.parseLine(lookForward ? this.rightPart() : this.leftPart());  
            
            counter = 0;
            total = null;
        }               
        forwardSearch = lookForward;
    }
    /* ������� ������������� ���������� ����������������� 
        ���������� ������ ������� VimAutoCompletionTool */
    this.isNewLoop = function (doc, word)
    {
        return !((srcDoc)
            &&(srcDoc.Path == doc.Path) 
            &&(srcLine == doc.SelStartLine) 
            &&(lastWord == word));         
    }
    /* ���������, �� ������� �� ������ ������ �� ���������� 
        ������� */
    this.assign = function (lIx)
    {
        return (srcDoc&&(0<=lIx)&&(lIx<srcDoc.LineCount));
    }
    /* ����� ��������� ������������ � ����������� ��� ��
        ����� ��������� ����� */
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
    /* ������ � ���������� ������ �������������� ���� ��� ��������� 
        �� ������� ������ */
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
    /* "���������" ���������� � �������� ��������� ������ �� �����
        � ��������� �� � ����������� � ��������, ������� ���������
        ���������� ����������� ��� ��������� �����*/
    this.parseLine = function (src)
    {
        var w = new Line(src);
        w.filter(pattern, true);
        w.setup(forwardSearch);
        return w;
    }
    /* ��������� ����������� ���������� ������������
        ������ ��������� ����� */
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
        // ����������� ������ ��� � Vim :-)       
        Status(counter 
            ? "C�����������: "+Math.abs(counter).toString()+(total?" �� " + total:"")
            : (total ? "�������� �����" : "������ �� ������")); 
    } 
    // ���������� ������� ������
    this.curLine = function ()
    {
        var src = "";
        if (this.assign(curLineIx)) {
            /* ��������� �������� ������ � ��� ��������� ��������,
                �� "��������" ��������, ��������� �� ����� �������� ����� */
            if (curLineIx == srcLine) {
                src = this.leftPart() + srcWord + this.rightPart();           
            }
            else {
                src = srcDoc.Range(curLineIx);
            }           
        }
        return src;
    }
    /* ����� �������� ������, ���������� �������� �����; 
        ���� �������� ����� �� ���������� */
    this.leftPart = function ()
    {
        return srcDoc.Range(srcLine, 0, srcLine, srcCol);
    }
    /* ������ �������� ������, ���������� �������� �����;
        ���� �������� ����� �� ���������� */
    this.rightPart = function ()
    {      
        return srcDoc.Range(
            srcLine, srcCol+lastWord.length,
            srcLine, srcDoc.LineLen(srcLine) - 1);
    }
}

/*
 * ��������� ������������� ������� 
 */
function Init(_) // ��������� ��������, ����� ��������� �� �������� � �������
{    
    try {
        var c = null;
        if (!(c = new ActiveXObject("OpenConf.CommonServices"))) {
            throw(true); // �������� ����������
        }        
        c.SetConfig(Configurator);
        SelfScript.AddNamedItem("CommonScripts", c, false);
    }
    catch (e) {
        Message("�� ���� ������� ������ OpenConf.CommonServices", mRedErr);    
        Message("������ " & SelfScript.Name & " �� ��������", mInformation);
        Scripts.UnLoad(SelfScript.Name);
    }
}

Init(); // ��� �������� ������� ��������� ��� �������������
