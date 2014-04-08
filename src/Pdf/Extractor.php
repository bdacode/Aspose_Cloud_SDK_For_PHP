<?php
/*
 * Extract various types of information from the document
 */
namespace Aspose\Cloud\Pdf;

use Aspose\Cloud\Common\AsposeApp;
use Aspose\Cloud\Common\Utils;
use Aspose\Cloud\Common\Product;
use Aspose\Cloud\Exception\AsposeCloudException as Exception;

class Extractor {

    public $fileName = '';

    public function __construct($fileName) {
        $this->fileName = $fileName;
    }

    /*
     * Gets number of images in a specified page
     * @param $pageNumber
     */
    public function getImageCount($pageNumber) {
        //check whether file is set or not
        if ($this->fileName == '')
            throw new Exception('No file name specified');

        $strURI = Product::$baseProductUri . '/pdf/' . $this->fileName . '/pages/' . $pageNumber . '/images';

        $signedURI = Utils::sign($strURI);

        $responseStream = Utils::processCommand($signedURI, 'GET', '', '');

        $json = json_decode($responseStream);

        return count($json->Images->List);
    }

    /*
     * Get the particular image from the specified page with the default image size
     * @param int $pageNumber
     * @param int $imageIndex
     * @param string $imageFormat
     */
    public function getImageDefaultSize($pageNumber, $imageIndex, $imageFormat) {
        //check whether file is set or not
        if ($this->fileName == '')
            throw new Exception('No file name specified');

        $strURI = Product::$baseProductUri . '/pdf/' . $this->fileName . '/pages/' . $pageNumber . '/images/' . $imageIndex . '?format=' . $imageFormat;

        $signedURI = Utils::sign($strURI);

        $responseStream = Utils::processCommand($signedURI, 'GET', '', '');

        $v_output = Utils::validateOutput($responseStream);

        if ($v_output === '') {
            $outputPath = AsposeApp::$outPutLocation . Utils::getFileName($this->fileName) . '_' . $imageIndex . '.' . $imageFormat;
            Utils::saveFile($responseStream, $outputPath);
            return $outputPath;
        }
        else
            return $v_output;
    }

    /*
     * Get the particular image from the specified page with the default image size
     * @param int $pageNumber
     * @param int $imageIndex
     * @param string $imageFormat
     * @param int $imageWidth
     * @param int $imageHeight
     */
    public function getImageCustomSize($pageNumber, $imageIndex, $imageFormat, $imageWidth, $imageHeight) {
        //check whether file is set or not
        if ($this->fileName == '')
            throw new Exception('No file name specified');

        $strURI = Product::$baseProductUri . '/pdf/' . $this->fileName . '/pages/' . $pageNumber . '/images/' . $imageIndex . '?format=' . $imageFormat . '&width=' . $imageWidth . '&height=' . $imageHeight;

        $signedURI = Utils::sign($strURI);
        $responseStream = Utils::processCommand($signedURI, 'GET', '', '');
        $v_output = Utils::validateOutput($responseStream);

        if ($v_output === '') {
            $outputPath = AsposeApp::$outPutLocation . Utils::getFileName($this->fileName) . '_' . $imageIndex . '.' . $imageFormat;
            Utils::saveFile($responseStream, $outputPath);
            return $outputPath;
        }
        else
            return $v_output;
    }

}