<?php
/*
 * This class contains features to work with text
 */
namespace Aspose\Cloud\Pdf;

use Aspose\Cloud\Common\AsposeApp;
use Aspose\Cloud\Common\Utils;
use Aspose\Cloud\Common\Product;
use Aspose\Cloud\Storage\Folder;
use Aspose\Cloud\Exception\AsposeCloudException as Exception;

class TextEditor {

    public $fileName = '';

    public function __construct($fileName) {
        $this->fileName = $fileName;
    }

    /*
     * Gets raw text from the whole PDF file or a specific page 
     */
    public function getText() {
        $parameters = func_get_args();

        //set parameter values
        if (count($parameters) > 0) {
            $pageNumber = $parameters[0];
        }
        //check whether file is set or not
        if ($this->fileName == '')
            throw new Exception('No file name specified');

        $strURI = Product::$baseProductUri . '/pdf/' . $this->fileName .
                ((isset($parameters[0])) ? '/pages/' . $pageNumber . '/TextItems' : '/TextItems');

        $signedURI = Utils::sign($strURI);

        $responseStream = Utils::processCommand($signedURI, 'GET', '', '');

        $json = json_decode($responseStream);

        $rawText = '';
        foreach ($json->TextItems->List as $textItem) {
            $rawText .= $textItem->Text;
        }
        return $rawText;
    }

    /*
     * Gets text items from the whole PDF file or a specific page 
     */
    public function getTextItems() {
        $parameters = func_get_args();

        //set parameter values
        if (count($parameters) == 1) {
            $pageNumber = $parameters[0];
        } else if (count($parameters) == 2) {
            $pageNumber = $parameters[0];
            $fragmentNumber = $parameters[1];
        }
        //check whether file is set or not
        if ($this->fileName == '')
            throw new Exception('No file name specified');

        $strURI = Product::$baseProductUri . '/pdf/' . $this->fileName;
        if (isset($parameters[0])) {
            $strURI .= '/pages/' . $pageNumber;
            if (isset($parameters[1])) {
                $strURI .= '/fragments/' . $fragmentNumber;
            }
        }
        $strURI .= '/TextItems';
        $signedURI = Utils::sign($strURI);

        $responseStream = Utils::processCommand($signedURI, 'GET', '', '');

        $json = json_decode($responseStream);

        return $json->TextItems->List;
    }

    /*
     * Gets count of the fragments from a particular page
     * $pageNumber
     */
    public function getFragmentCount($pageNumber) {
        //check whether file is set or not
        if ($this->fileName == '')
            throw new Exception('No file name specified');

        $strURI = Product::$baseProductUri . '/pdf/' . $this->fileName . '/pages/' . $pageNumber . '/fragments';

        $signedURI = Utils::sign($strURI);

        $responseStream = Utils::processCommand($signedURI, 'GET', '', '');

        $json = json_decode($responseStream);

        return count($json->TextItems->List);
    }

    /*
     * Gets count of the segments in a fragment
     * @param number $pageNumber
     * @param number $fragmentNumber
     */
    public function getSegmentCount($pageNumber = '', $fragmentNumber = '') {
        //check whether file is set or not
        if ($this->fileName == '')
            throw new Exception('No file name specified');

        if ($pageNumber == '')
            throw new Exception('page number not specified');

        if ($fragmentNumber == '')
            throw new Exception('fragment number not specified');


        $strURI = Product::$baseProductUri . '/pdf/' . $this->fileName . '/pages/' . $pageNumber . '/fragments/' . $fragmentNumber;

        $signedURI = Utils::sign($strURI);

        $responseStream = Utils::processCommand($signedURI, 'GET', '', '');

        $json = json_decode($responseStream);

        return count($json->TextItems->List);
    }

    /*
     * Gets TextFormat of a particular Fragment
     * $pageNumber
     * $fragmentNumber
     */
    public function getTextFormat() {
        $args = func_get_args();
        if ($this->fileName == '')
            throw new Exception('No file name specified');
        if (count($args) == 2) {
            $pageNumber = $args[0];
            $fragmentNumber = $args[1];
        }
        if (count($args) == 3) {
            $pageNumber = $args[0];
            $fragmentNumber = $args[1];
            $segamentNumber = $args[2];
        }
        $strURI = Product::$baseProductUri . '/pdf/' . $this->fileName . '/pages/' . $pageNumber .
                '/fragments/' . $fragmentNumber;
        if (isset($segamentNumber)) {
            $strURI .= '/segments/' . '/textformat';
        } else {
            $strURI .= '/textformat';
        }
        $signedURI = Utils::sign($strURI);

        $responseStream = Utils::processCommand($signedURI, 'GET', '', '');

        $json = json_decode($responseStream);

        return $json->TextFormat;
    }

    /*
     * Replaces all instances of old text with new text in a PDF file or a particular page
     * @param Array $fieldsArray
     */
    public function replaceMultipleText($fieldsArray) {
        //check whether file is set or not
        if ($this->fileName == '')
            throw new Exception('No file name specified');

        //Build JSON to post
        $json = json_encode($fieldsArray);

        //Build URI to replace text
        $strURI = Product::$baseProductUri . '/pdf/' . $this->fileName . '/replaceTextList';

        $signedURI = Utils::sign($strURI);

        $responseStream = Utils::processCommand($signedURI, 'POST', 'json', $json);

        $v_output = Utils::validateOutput($responseStream);

        if ($v_output === '') {
            //Save doc on server
            $folder = new Folder();
            $outputStream = $folder->GetFile($this->fileName);
            $outputPath = AsposeApp::$outPutLocation . $this->fileName;
            Utils::saveFile($outputStream, $outputPath);
            return $outputPath;
        }
        else
            return $v_output;
    }

    /*
     * Replaces all instances of old text with new text in a PDF file or a particular page
     * @param string $oldText
     * @param string $newText
     */
    public function replaceText() {
        $parameters = func_get_args();

        //set parameter values
        if (count($parameters) == 3) {
            $oldText = $parameters[0];
            $newText = $parameters[1];
            $isRegularExpression = $parameters[2];
        } else if (count($parameters) == 4) {
            $oldText = $parameters[0];
            $newText = $parameters[1];
            $isRegularExpression = $parameters[2];
            $pageNumber = $parameters[3];
        }
        else
            throw new Exception('Invalid number of arguments');
        //check whether file is set or not
        if ($this->fileName == '')
            throw new Exception('No file name specified');

        //Build JSON to post
        $fieldsArray = array('OldValue' => $oldText, 'NewValue' => $newText, 'Regex' => $isRegularExpression);
        $json = json_encode($fieldsArray);

        //Build URI to replace text
        $strURI = Product::$baseProductUri . '/pdf/' . $this->fileName . ((isset($parameters[3])) ? '/pages/' . $pageNumber : '') .
                '/replaceText';

        $signedURI = Utils::sign($strURI);

        $responseStream = Utils::processCommand($signedURI, 'POST', 'json', $json);

        $v_output = Utils::validateOutput($responseStream);

        if ($v_output === '') {
            //Save doc on server
            $folder = new Folder();
            $outputStream = $folder->GetFile($this->fileName);
            $outputPath = AsposeApp::$outPutLocation . $this->fileName;
            Utils::saveFile($outputStream, $outputPath);
            return $outputPath;
        }
        else
            return $v_output;
    }
}