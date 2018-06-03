<?php

/* 
 * A collection of simple tools for analysing 
 * .PDF, .DOCX, .DOC, .RTF and .TXT docs. 
 * 
 *  Copyright (C) 2016-2017
 *    Joseph Blurton (http://github.com/joeblurton)
 *    And other contributors (see attrib below)
 *  
 *  Version 1.0.2
 *
 *  This program is free software: you can redistribute it and/or modify
 *  it under the terms of the GNU General Public License as published by
 *  the Free Software Foundation, either version 3 of the License, or
 *  (at your option) any later version.
 *
 *  This program is distributed in the hope that it will be useful,
 *  but WITHOUT ANY WARRANTY; without even the implied warranty of
 *  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 *  GNU General Public License for more details.
 *
 *  You should have received a copy of the GNU General Public License
 *  along with this program.  If not, see <http://www.gnu.org/licenses/>.
 *
 * ATTRIBUTIONS
 *
 * PageCount_PDF and 
 * PageCount_DOCX by Whiteflash
 * http://stackoverflow.com/questions/5540886/extract-text-from-doc-and-docx/
 *
 * Paragraph tweak by JoshB
 * http://stackoverflow.com/questions/5607594/find-linebreaks-in-a-docx-file-using-php
 * 
 * read_word_doc by
 * Davinder Singh
 * http://stackoverflow.com/questions/7358637/reading-doc-file-in-php
 *
 * Jonny 5's simple word splitter
 * http://php.net/manual/en/function.str-word-count.php#107363
 * 
 * Line Count method by K2xL
 * http://stackoverflow.com/questions/7955402/count-lines-in-a-posted-string
 *
 * RTFTOOLS by
 * Christian Vigh
 * https://github.com/christian-vigh-phpclasses/RtfTools
 *
 * PDF Parser by
 * Smalot GPL 3
 * https://github.com/smalot/pdfparser
 */

class DocCounter {
    
    // Class Variables
    public $pdfTempFile = '/tmp/temp.pdf';
    private $file;
    private $filetype;
    
    // Set file
    public function setFile($filename)
    {
        $this->file = $filename;
        $this->filetype = pathinfo($this->file, PATHINFO_EXTENSION);
    }
    
    // Get file
    public function getFile()
    {
        return $this->file;
    }
    
    // Get file information object
    public function getInfo()
    {
        // Function variables
        $ft = $this->filetype;
        
        // Let's construct our info response object
        $obj = new stdClass();
        $obj->format = $ft;
        $obj->wordCount = null;
        $obj->lineCount = null;
        $obj->pageCount = null;
        
        // Let's set our function calls based on filetype
        switch($ft)
        {
            case "doc":
                $doc = $this->read_doc_file();
                $obj->wordCount = $this->str_word_count_utf8($doc);
                $obj->lineCount = $this->lineCount($doc);
                $obj->pageCount = $this->pageCount($doc);
                break;
            case "docx":
                $obj->wordCount = $this->str_word_count_utf8($this->docx2text());
                $obj->lineCount = $this->lineCount($this->docx2text());
                $obj->pageCount = $this->PageCount_DOCX();
                break;
            case "pdf":
                $obj->wordCount = $this->str_word_count_utf8($this->pdf2text());
                $obj->lineCount = $this->lineCount($this->pdf2text());
                $obj->pageCount = $this->PageCount_PDF();
                break;
            case "txt":
                $textContents = file_get_contents($this->file);
                $obj->wordCount = $this->str_word_count_utf8($textContents);
                $obj->lineCount = $this->lineCount($textContents);
                $obj->pageCount = $this->pageCount($textContents);
                break;
            case "rtf":
                $textContents = $this->rtf2text();
                $obj->wordCount = $this->str_word_count_utf8($textContents);
                $obj->lineCount = $this->lineCount($textContents);
                $obj->pageCount = $this->pageCount($textContents);
                break;
            default:
                $obj->wordCount = "unsupported file format";
                $obj->lineCount = "unsupported file format";
                $obj->pageCount = "unsupported file format";
        }
        
        return $obj;
    }
    
    // Convert: Word.doc to Text String
    function read_doc_file() {
        
        $f = $this->file;
         if(file_exists($f))
        {
            if(($fh = fopen($f, 'r')) !== false ) 
            {
               $headers = fread($fh, 0xA00);

               // 1 = (ord(n)*1) ; Document has from 0 to 255 characters
               $n1 = ( ord($headers[0x21C]) - 1 );

               // 1 = ((ord(n)-8)*256) ; Document has from 256 to 63743 characters
               $n2 = ( ( ord($headers[0x21D]) - 8 ) * 256 );

               // 1 = ((ord(n)*256)*256) ; Document has from 63744 to 16775423 characters
               $n3 = ( ( ord($headers[0x21E]) * 256 ) * 256 );

               // 1 = (((ord(n)*256)*256)*256) ; Document has from 16775424 to 4294965504 characters
               $n4 = ( ( ( ord($headers[0x21F]) * 256 ) * 256 ) * 256 );

               // Total length of text in the document
               $textLength = ($n1 + $n2 + $n3 + $n4);

               $extracted_plaintext = fread($fh, $textLength);
                $extracted_plaintext = mb_convert_encoding($extracted_plaintext,'UTF-8');
               // simple print character stream without new lines
               //echo $extracted_plaintext;

               // if you want to see your paragraphs in a new line, do this
               return nl2br($extracted_plaintext);
               // need more spacing after each paragraph use another nl2br
            }
        }
    }
    // Jonny 5's simple word splitter
    function str_word_count_utf8($str) {
        //return count(preg_split('~[^\p{L}\p{N}\']+~u',$str));
	return str_word_count($str);
    }
    // Convert: Word.docx to Text String
    function docx2text()
    {
        return $this->readZippedXML($this->file, "word/document.xml");
    }

    function readZippedXML($archiveFile, $dataFile)
    {
        // Create new ZIP archive
        $zip = new ZipArchive;
        
        // set absolute path
        $f = $archiveFile;

        // Open received archive file
        if (true === $zip->open($f)) {
            // If done, search for the data file in the archive
            if (($index = $zip->locateName($dataFile)) !== false) {
                // If found, read it to the string
                $data = $zip->getFromIndex($index);
                // Close archive file
                $zip->close();
                // Load XML from a string
                // Skip errors and warnings
                $xml = new DOMDocument();
                $xml->loadXML($data, LIBXML_NOENT | LIBXML_XINCLUDE | LIBXML_NOERROR | LIBXML_NOWARNING);
                
                $xmldata = $xml->saveXML();
                // Newline Replacement
                $xmldata = str_replace("</w:p>", "\r\n", $xmldata);
                // Return data without XML formatting tags
                return strip_tags($xmldata);
            }
            $zip->close();
        }

        // In case of failure return empty string
        return "";
    }
    
    // Convert: Word.doc to Text String
    function read_doc()
    {
        $f = $this->file;
        $fileHandle = fopen($f, "r");
        $line = @fread($fileHandle, filesize($this->file));   
        $lines = explode(chr(0x0D),$line);
        $outtext = "";
        foreach($lines as $thisline)
          {
            $pos = strpos($thisline, chr(0x00));
            if (($pos !== FALSE)||(strlen($thisline)==0))
              {
              } else {
                $outtext .= $thisline." ";
              }
          }
        $outtext = preg_replace("/[^a-zA-Z0-9\s\,\.\-\n\r\t@\/\_\(\)]/","",$outtext);
        return $outtext;
    }
    
    // Extract text from RTF doc
    function rtf2text()
    {
        $f = $this->file;
        
        if (file_exists($f)) {
            $input_lines = file_get_contents($this->file);
            
            preg_match_all("/\\\\([a-z]{1,32})(-?\d{1,10})?[ ]?|\\\\'([0-9a-f]{2})|\\\\([^a-z])|([{}])|[\r\n]+|(.)/i", $input_lines, $output_array);

            $stack = [];
            $ignorable = false;
            $ucskip = 1;
            $curskip = 0;
            $out = [];
            $matches = count($output_array[0]);

            $destinations = [
      'aftncn','aftnsep','aftnsepc','annotation','atnauthor','atndate','atnicn','atnid',
      'atnparent','atnref','atntime','atrfend','atrfstart','author','background',
      'bkmkend','bkmkstart','blipuid','buptim','category','colorschememapping',
      'colortbl','comment','company','creatim','datafield','datastore','defchp','defpap',
      'do','doccomm','docvar','dptxbxtext','ebcend','ebcstart','factoidname','falt',
      'fchars','ffdeftext','ffentrymcr','ffexitmcr','ffformat','ffhelptext','ffl',
      'ffname','ffstattext','field','file','filetbl','fldinst','fldrslt','fldtype',
      'fname','fontemb','fontfile','fonttbl','footer','footerf','footerl','footerr',
      'footnote','formfield','ftncn','ftnsep','ftnsepc','g','generator','gridtbl',
      'header','headerf','headerl','headerr','hl','hlfr','hlinkbase','hlloc','hlsrc',
      'hsv','htmltag','info','keycode','keywords','latentstyles','lchars','levelnumbers',
      'leveltext','lfolevel','linkval','list','listlevel','listname','listoverride',
      'listoverridetable','listpicture','liststylename','listtable','listtext',
      'lsdlockedexcept','macc','maccPr','mailmerge','maln','malnScr','manager','margPr',
      'mbar','mbarPr','mbaseJc','mbegChr','mborderBox','mborderBoxPr','mbox','mboxPr',
      'mchr','mcount','mctrlPr','md','mdeg','mdegHide','mden','mdiff','mdPr','me',
      'mendChr','meqArr','meqArrPr','mf','mfName','mfPr','mfunc','mfuncPr','mgroupChr',
      'mgroupChrPr','mgrow','mhideBot','mhideLeft','mhideRight','mhideTop','mhtmltag',
      'mlim','mlimloc','mlimlow','mlimlowPr','mlimupp','mlimuppPr','mm','mmaddfieldname',
      'mmath','mmathPict','mmathPr','mmaxdist','mmc','mmcJc','mmconnectstr',
      'mmconnectstrdata','mmcPr','mmcs','mmdatasource','mmheadersource','mmmailsubject',
      'mmodso','mmodsofilter','mmodsofldmpdata','mmodsomappedname','mmodsoname',
      'mmodsorecipdata','mmodsosort','mmodsosrc','mmodsotable','mmodsoudl',
      'mmodsoudldata','mmodsouniquetag','mmPr','mmquery','mmr','mnary','mnaryPr',
      'mnoBreak','mnum','mobjDist','moMath','moMathPara','moMathParaPr','mopEmu',
      'mphant','mphantPr','mplcHide','mpos','mr','mrad','mradPr','mrPr','msepChr',
      'mshow','mshp','msPre','msPrePr','msSub','msSubPr','msSubSup','msSubSupPr','msSup',
      'msSupPr','mstrikeBLTR','mstrikeH','mstrikeTLBR','mstrikeV','msub','msubHide',
      'msup','msupHide','mtransp','mtype','mvertJc','mvfmf','mvfml','mvtof','mvtol',
      'mzeroAsc','mzeroDesc','mzeroWid','nesttableprops','nextfile','nonesttables',
      'objalias','objclass','objdata','object','objname','objsect','objtime','oldcprops',
      'oldpprops','oldsprops','oldtprops','oleclsid','operator','panose','password',
      'passwordhash','pgp','pgptbl','picprop','pict','pn','pnseclvl','pntext','pntxta',
      'pntxtb','printim','private','propname','protend','protstart','protusertbl','pxe',
      'result','revtbl','revtim','rsidtbl','rxe','shp','shpgrp','shpinst',
      'shppict','shprslt','shptxt','sn','sp','staticval','stylesheet','subject','sv',
      'svb','tc','template','themedata','title','txe','ud','upr','userprops',
      'wgrffmtfilter','windowcaption','writereservation','writereservhash','xe','xform',
      'xmlattrname','xmlattrvalue','xmlclose','xmlname','xmlnstbl',
      'xmlopen'];
	  
            $specialchars = [
      'par' => "\n",
      'sect' => "\n\n",
      'page' => "\n\n",
      'line' => "\n",
      'tab' => "\t",
      'emdash' => "\u2014",
      'endash' => "\u2013",
      'emspace' => "\u2003",
      'enspace' => "\u2002",
      'qmspace' => "\u2005",
      'bullet' => "\u2022",
      'lquote' => "\u2018",
      'rquote' => "\u2019",
      'ldblquote' => "\201C",
      'rdblquote' => "\u201D"];
            
            for ($i = 0; $i < $matches; $i++) {
                $word = $output_array[1][$i];
                $arg = $output_array[2][$i];
                $hex = $output_array[3][$i];
                $char = $output_array[4][$i];
                $brace = $output_array[5][$i];
                $tchar = $output_array[6][$i];
	
                // braces
                if (!empty($brace)) {
                    $curskip = 0;

                    if ($brace == '{') {
                        array_push($stack, [$ucskip, $ignorable]);
                    } else if ($brace == '}') {
                        list($ucskip, $ignorable) = array_pop($stack);
                    }
                }
                // not a letter
                else if (!empty($char)) {
                    $curskip = 0;

                    if ($char == '~') {
                        if (!$ignorable) {
                            $out[] = "\xA0";
                        }
                    } else if (in_array($char, ['{','}','\\'])) {
                        if (!$ignorable) {
                            $out[] = $char;
                        }
                    } else if ($char == '*') {
                        $ignorable = true;
                    }
                }
                //
                else if (!empty($word)) {
                    $curskip = 0;

                    if (in_array($word, $destinations)) {
                        $ignorable = true;
                    } else if ($ignorable) {

                    } else if (!empty($specialchars[$word])) {
                        $out[] = $specialchars[$word];
                    } else if ($word == 'uc') {
                        $ucskip = (int)$arg;
                    } else if ($word == 'u') {
                        $c = (int)$arg;
                        if ($c < 0) {
                            $c += 0x10000;
                        }
                        if ($c > 127) {
                            $out[] = chr($c);
                        } else {
                            $out[] = chr($c);
                        }

                        $curskip = $ucskip;
                    }
                }
                //
                else if (!empty($hex)) {
                    if ($curskip > 0) {
                        $curskip -= 1;
                    } else if (!$ignorable) {
                        $c = hexdec($hex);
                        if ($c > 127) {
                            $out[] = chr($c);
                        } else {
                            $out[] = chr($c);
                        }
                    }
                }
                //
                else if (!empty($tchar)) {
                    if ($curskip > 0) {
                        $curskip -= 1;
                    } else if (!$ignorable) {
                        $out[] = $tchar;
                    }
                }
            }

            return join('', $out);
        }
        
        return null;
    }
    
    // Convert: Adobe.pdf to Text String
    function pdf2text()
    {
        //absolute path for file
        $f = $this->file;
        
        if (file_exists($f)) {
            include('vendor/autoload.php');
            $parser = new \Smalot\PdfParser\Parser();
            $pdf = $parser->parseFile($f);
            $text = $pdf->getText();
            return $text;
        }
        
        return null;
    }
    
    // Page Count: DOCX using XML Metadata
    function PageCount_DOCX()
    {
        $pageCount = 0;

        $zip = new ZipArchive();
        
        $f = $this->file;

        if($zip->open($f) === true) {
            if(($index = $zip->locateName('docProps/app.xml')) !== false)  {
                $data = $zip->getFromIndex($index);
                $zip->close();
                $xml = new SimpleXMLElement($data);
                $pageCount = $xml->Pages;
            }
        }

        return intval($pageCount);
    }

    // Page Count: PDF using FPDF and FPDI 
    function PageCount_PDF()
    {
        //absolute path for file
        $f = $this->file;
        $pageCount = 0;
        if (file_exists($f)) {
            require_once('lib/fpdf/fpdf.php');
            require_once('lib/fpdi/fpdi.php');
            $pdf = new FPDI();
            $pageCount = $pdf->setSourceFile($f);        // returns page count
        }
        return $pageCount;
    }
    
    // Page Count: General
    function pageCount($text)
    {
        require_once('lib/fpdf/fpdf.php');

        $pdf = new FPDF();
        $pdf->AddPage();
        $pdf->SetFont('Times','',12);
        $pdf->MultiCell(0,5,$text);
        //$pdf->Output();
        $filename = $this->pdfTempFile;
        $pdf->Output($filename,'F');
        
        require_once('lib/fpdi/fpdi.php');
        $pdf = new FPDI();
        $pageCount = $pdf->setSourceFile($filename);
        
        unlink($filename);
        return $pageCount;
    }
    
    // Line Count: General
    function lineCount($text)
    {
        $lines_arr = preg_split('/\n|\r/',$text);
        $num_newlines = count($lines_arr); 
        return $num_newlines;
    }
}


?>
