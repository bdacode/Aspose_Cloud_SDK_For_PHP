<?php
/*
 * converts pages or document into different formats
 */
namespace Aspose\Cloud\Slides;

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

        $this->saveFormat = 'PPT';
    }

    /*
     * Saves a particular slide into various formats with specified width and height
     * @param string $slideNumber
     * @param string $imageFormat
     */
    public function convertToImage($slideNumber, $imageFormat) {
        try {
            //check whether file is set or not
            if ($this->fileName == '')
                throw new Exception('No file name specified');

            $strURI = Product::$baseProductUri . '/slides/' . $this->fileName . '/slides/' . $slideNumber . '?format=' . $imageFormat;

            $signedURI = Utils::sign($strURI);

            $responseStream = Utils::processCommand($signedURI, 'GET', '', '');
            $v_output = Utils::validateOutput($responseStream);
            if ($v_output == '') {
                $outputPath = AsposeApp::$outPutLocation . Utils::getFileName($this->fileName) . '.' . $imageFormat;
                Utils::saveFile($responseStream, $outputPath);
                return $outputPath;
            } else {
                return $v_output;
            }
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    /*
     * Saves a particular slide into various formats with specified width and height
     * @param string $slideNumber
     * @param string $imageFormat
     * @param string $width
     * @param string $height
     */
    public function convertToImagebySize($slideNumber, $imageFormat, $width, $height) {
        try {
            //check whether file is set or not
            if ($this->fileName == '')
                throw new Exception('No file name specified');

            $strURI = Product::$baseProductUri . '/slides/' . $this->fileName . '/slides/' . $slideNumber . '?format=' . $imageFormat . '&width=' . $width . '&height=' . $height;

            $signedURI = Utils::sign($strURI);

            $responseStream = Utils::processCommand($signedURI, 'GET', '', '');
            $v_output = Utils::validateOutput($responseStream);
            if ($v_output == '') {
                $outputPath = AsposeApp::$outPutLocation . 'output.' . $imageFormat;
                Utils::saveFile($responseStream, $outputPath);
                return $outputPath;
            } else {
                return $v_output;
            }
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    /*
     * convert a document to SaveFormat
     */
    public function convert() {
        try {
            //check whether file is set or not
            if ($this->fileName == '')
                throw new Exception('No file name specified');

            $strURI = Product::$baseProductUri . '/slides/' . $this->fileName . '?format=' . $this->saveFormat;

            $signedURI = Utils::sign($strURI);
            $responseStream = Utils::processCommand($signedURI, 'GET', '', '');
            
            $v_output = Utils::validateOutput($responseStream);

            if ($v_output === '') {
                if ($this->saveFormat == 'html') {
                    $save_format = 'zip';
                } else {
                    $save_format = $this->saveFormat;
                }
                $outputPath =  AsposeApp::$outPutLocation . Utils::getFileName($this->fileName) . '.' . $save_format;
                Utils::saveFile($responseStream,$outputPath);
                return $outputPath;
            } else {
                return $v_output;
            }
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

}