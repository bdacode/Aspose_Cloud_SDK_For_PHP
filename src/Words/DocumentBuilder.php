<?php
/*
 * Deals with Word document builder aspects
 */
namespace Aspose\Cloud\Words;

use Aspose\Cloud\Common\AsposeApp;
use Aspose\Cloud\Common\Utils;
use Aspose\Cloud\Common\Product;
use Aspose\Cloud\Storage\Folder;
use Aspose\Cloud\Exception\AsposeCloudException as Exception;

class DocumentBuilder {


    /*
     * Remove watermark from document
     */
    public function removeWatermark($fileName) {
        //check whether files are set or not
        if ($fileName == '')
            throw new Exception('File not specified');

        //build URI to insert watermark image
        $strURI = Product::$baseProductUri . '/words/' . $fileName .
            '/watermark';

        //sign URI
        $signedURI = Utils::sign($strURI);

        $responseStream = Utils::processCommand($signedURI, 'DELETE', '', '');

        $v_output = Utils::validateOutput($responseStream);

        if ($v_output === '') {
            //Save doc on server
            $folder = new Folder();
            $outputStream = $folder->GetFile($fileName);
            $outputPath = AsposeApp::$outPutLocation . $fileName;
            Utils::saveFile($outputStream, $outputPath);
            return $outputPath;
        }
        else
            return $v_output;
    }

    /*
     * Inserts water mark text into the document.
     * @param string $fileName 
     * @param string $text
     * @param string $rotationAngle 
     */
    public function insertWatermarkText($fileName, $text, $rotationAngle) {
        //check whether files are set or not
        if ($fileName == '')
            throw new Exception('File not specified');

        //Build JSON to post
        $fieldsArray = array('Text' => $text, 'RotationAngle' => $rotationAngle);
        $json = json_encode($fieldsArray);

        //build URI to insert watermark text
        $strURI = Product::$baseProductUri . '/words/' . $fileName . '/watermark/insertText';

        //sign URI
        $signedURI = Utils::sign($strURI);

        $responseStream = Utils::processCommand($signedURI, 'POST', 'json', $json);

        $v_output = Utils::validateOutput($responseStream);

        if ($v_output === '') {
            //Save docs on server
            $folder = new Folder();
            $outputStream = $folder->GetFile($fileName);
            $outputPath = AsposeApp::$outPutLocation . $fileName;
            Utils::saveFile($outputStream, $outputPath);
            return $outputPath;
        }
        else
            return $v_output;
    }

    /*
     * Inserts water mark image into the document.
     * @param string $fileName 
     * @param string $imageFile
     * @param string $rotationAngle 
     */
    public function insertWatermarkImage($fileName, $imageFile, $rotationAngle) {
        //check whether files are set or not
        if ($fileName == '')
            throw new Exception('File not specified');

        //build URI to insert watermark image
        $strURI = Product::$baseProductUri . '/words/' . $fileName .
                '/watermark/insertImage?imageFile=' . $imageFile . '&rotationAngle=' . $rotationAngle;

        //sign URI
        $signedURI = Utils::sign($strURI);

        $responseStream = Utils::processCommand($signedURI, 'POST', 'json', '');

        $v_output = Utils::validateOutput($responseStream);

        if ($v_output === '') {
            //Save doc on server
            $folder = new Folder();
            $outputStream = $folder->GetFile($fileName);
            $outputPath = AsposeApp::$outPutLocation . $fileName;
            Utils::saveFile($outputStream, $outputPath);
            return $outputPath;
        }
        else
            return $v_output;
    }

    /*
     * Replace a text with the new value in the document
     * @param string $fileName 
     * @param string $oldValue
     * @param string $newValue 
     * @param string $isMatchCase
     * @param string $isMatchWholeWord
     */
    public function replaceText($fileName, $oldValue, $newValue, $isMatchCase, $isMatchWholeWord) {
        //check whether files are set or not
        if ($fileName == '')
            throw new Exception('File not specified');

        //Build JSON to post
        $fieldsArray = array('OldValue' => $oldValue, 'NewValue' => $newValue,
            'IsMatchCase' => $isMatchCase, 'IsMatchWholeWord' => $isMatchWholeWord);
        $json = json_encode($fieldsArray);

        //build URI to replace text
        $strURI = Product::$baseProductUri . '/words/' . $fileName . '/replaceText';

        //sign URI
        $signedURI = Utils::sign($strURI);

        $responseStream = Utils::processCommand($signedURI, 'POST', 'json', $json);

        $v_output = Utils::validateOutput($responseStream);

        if ($v_output === '') {
            //Save docs on server
            $folder = new Folder();
            $outputStream = $folder->GetFile($fileName);
            $outputPath = AsposeApp::$outPutLocation . $fileName;
            Utils::saveFile($outputStream, $outputPath);
            return $outputPath;
        }
        else
            return $v_output;
    }

}