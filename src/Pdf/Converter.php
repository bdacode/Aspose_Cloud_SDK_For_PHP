<?php
/*
 * converts pages or document into different formats
 */
namespace Aspose\Cloud\Pdf;

use Aspose\Cloud\Common\AsposeApp;
use Aspose\Cloud\Common\Utils;
use Aspose\Cloud\Common\Product;
use Aspose\Cloud\Exception\AsposeCloudException as Exception;

class Converter {

    public $fileName = '';
    public $saveFormat = '';

    public function __construct($fileName) {
        $this->fileName = $fileName;

        $this->saveFormat = 'Pdf';
    }

    /*
     * convert a particular page to image with specified size
     * @param string $pageNumber
     * @param string $imageFormat
     * @param string $width
     * @param string $height
     */
    public function convertToImagebySize($pageNumber, $imageFormat, $width, $height) {
        try {
            //check whether file is set or not
            if ($this->fileName == '')
                throw new Exception('No file name specified');

            $strURI = Product::$baseProductUri . '/pdf/' . $this->fileName . '/pages/' . $pageNumber . '?format=' . $imageFormat . '&width=' . $width . '&height=' . $height;

            $signedURI = Utils::sign($strURI);

            $responseStream = Utils::processCommand($signedURI, 'GET', '', '');

            $v_output = Utils::validateOutput($responseStream);

            if ($v_output === '') {
                $outputPath = AsposeApp::$outPutLocation . Utils::getFileName($this->fileName) . '_' . $pageNumber . '.' . $imageFormat;
                Utils::saveFile($responseStream, $outputPath);
                return $outputPath;
            }
            else
                return $v_output;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    /*
     * convert a particular page to image with default size
     * @param string $pageNumber
     * @param string $imageFormat
     */
    public function convertToImage($pageNumber, $imageFormat) {
        try {
            //check whether file is set or not
            if ($this->fileName == '')
                throw new Exception('No file name specified');

            $strURI = Product::$baseProductUri . '/pdf/' . $this->fileName . '/pages/' . $pageNumber . '?format=' . $imageFormat;

            $signedURI = Utils::sign($strURI);

            $responseStream = Utils::processCommand($signedURI, 'GET', '', '');

            $v_output = Utils::validateOutput($responseStream);

            if ($v_output === '') {
                $outputPath = AsposeApp::$outPutLocation . Utils::getFileName($this->fileName) . '_' . $pageNumber . '.' . $imageFormat;
                Utils::saveFile($responseStream, $outputPath);
                return $outputPath;
            }
            else
                return $v_output;
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

            $strURI = Product::$baseProductUri . '/pdf/' . $this->fileName . '?format=' . $this->saveFormat;

            $signedURI = Utils::sign($strURI);

            $responseStream = Utils::processCommand($signedURI, 'GET', '', '');

            $v_output = Utils::validateOutput($responseStream);

            if ($v_output === '') {
                if ($this->saveFormat == 'html') {
                    $saveFormat = 'zip';
                } else {
                    $saveFormat = $this->saveFormat;
                }

                $outputPath = Utils::saveFile($responseStream, AsposeApp::$outPutLocation . Utils::getFileName($this->fileName) . '.' . $saveFormat);
                return $outputPath;
            } else {
                return $v_output;
            }
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    /*
     * Convert PDF to different file format without using storage
     * $param string $inputFile
     * @param string $outputFilename
     * @param string $outputFormat
     */
    public function convertLocalFile($inputFile = '', $outputFilename = '', $outputFormat = '') {
        try {
            //check whether file is set or not
            if ($inputFile == '')
                throw new Exception('No file name specified');

            if ($outputFormat == '')
                throw new Exception('output format not specified');


            $strURI = Product::$baseProductUri . '/pdf/convert?format=' . $outputFormat;

            if (!file_exists($inputFile)) {
                throw new Exception('input file doesnt exist.');
            }


            $signedURI = Utils::sign($strURI);

            $responseStream = Utils::uploadFileBinary($signedURI, $inputFile, 'xml');

            $v_output = Utils::validateOutput($responseStream);

            if ($v_output === '') {
                if ($outputFormat == 'html') {
                    $saveFormat = 'zip';
                } else {
                    $saveFormat = $outputFormat;
                }

                if ($outputFilename == '') {
                    $outputFilename = Utils::getFileName($inputFile) . '.' . $saveFormat;
                }
                $outputPath = AsposeApp::$outPutLocation . $outputFilename;
                Utils::saveFile($responseStream, $outputPath);
                return $outputPath;
            }
            else
                return $v_output;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

}