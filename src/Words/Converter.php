<?php
/*
 * converts pages or document into different formats
 */
namespace Aspose\Cloud\Words;

use Aspose\Cloud\Common\AsposeApp;
use Aspose\Cloud\Common\Utils;
use Aspose\Cloud\Common\Product;
use Aspose\Cloud\Exception\AsposeCloudException as Exception;

class Converter {

    public $fileName = '';
    public $saveFormat = '';

    public function __construct($fileName) {
        //set default values
        $this->fileName = $fileName;

        $this->saveFormat = 'Doc';
    }

    /*
     * convert a document to SaveFormat
     */
    public function convert() {
        try {
            //check whether file is set or not
            if ($this->fileName == '')
                throw new Exception('No file name specified');

            //build URI
            $strURI = Product::$baseProductUri . '/words/' . $this->fileName . '?format=' . $this->saveFormat;

            //sign URI
            $signedURI = Utils::sign($strURI);

            $responseStream = Utils::processCommand($signedURI, 'GET', '', '');

            $v_output = Utils::validateOutput($responseStream);

            if ($v_output === '') {
                if ($this->saveFormat == 'html') {
                    $save_format = 'zip';
                } else {
                    $save_format = $this->saveFormat;
                }
                $outputPath = AsposeApp::$outPutLocation . Utils::getFileName($this->fileName) . '.' . $save_format;
                Utils::saveFile($responseStream, $outputPath);
                return $outputPath;
            } else {
                return $v_output;
            }
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    public function convertLocalFile($inputPath, $outputPath, $outputFormat) {
        try {
            $str_uri = Product::$baseProductUri . '/words/convert?format=' . $outputFormat;
            $signed_uri = Utils::sign($str_uri);
            $responseStream = Utils::uploadFileBinary($signed_uri, $inputPath, 'xml');

            $v_output = Utils::validateOutput($responseStream);

            if ($v_output === '') {
                if ($outputFormat == 'html') {
                    $saveFormat = 'zip';
                } else {
                    $saveFormat = $outputFormat;
                }

                if ($outputPath == '') {
                    $outputFilename = Utils::getFileName($inputPath) . '.' . $saveFormat;
                }

                Utils::saveFile($responseStream, AsposeApp::$outPutLocation . $outputFilename);
                return $outputFilename;
            }
            else
                return $v_output;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

}