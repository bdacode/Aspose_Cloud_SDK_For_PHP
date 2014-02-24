<?php
/**
 * Created by PhpStorm.
 * User: AssadMahmood
 * Date: 2/24/14
 * Time: 2:59 PM
 */

namespace Aspose\Cloud\Imaging;

use Aspose\Cloud\Common\AsposeApp;
use Aspose\Cloud\Common\Utils;
use Aspose\Cloud\Common\Product;
use Aspose\Cloud\Storage\Folder;
use Aspose\Cloud\Exception\AsposeCloudException as Exception;

class Image {

    public function __construct($fileName){
        $this->fileName = $fileName;
    }

    public function convertTiffToFax(){

        try {
            //check whether file is set or not
            if ($this->fileName == '')
                throw new Exception('No file name specified');

            $strURI = Product::$baseProductUri . '/imaging/tiff/' . $this->fileName . '/toFax';

            $signedURI = Utils::sign($strURI);

            $responseStream = Utils::processCommand($signedURI, 'GET', '', '');

            $v_output = Utils::validateOutput($responseStream);

            if ($v_output === '') {
                $outputPath = AsposeApp::$outPutLocation . $this->fileName;
                Utils::saveFile($responseStream, $outputPath);
                return $outputPath;
            }
            else
                return $v_output;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }

    }

    public function appendTiff($appendFile=""){

        try {
            //check whether file is set or not
            if ($this->fileName == '' || $appendFile == '')
                throw new Exception('No file name specified');

            $strURI = Product::$baseProductUri . '/imaging/tiff/' . $this->fileName . '/appendTiff?appendFile=' . $appendFile;

            $signedURI = Utils::sign($strURI);

            $responseStream = Utils::processCommand($signedURI, 'POST', '', '');

            $json = json_decode($responseStream);

            if ($json->Status == 'OK') {
                $folder = new Folder();
                $outputStream = $folder->getFile($this->fileName);
                $outputPath = AsposeApp::$outPutLocation . $this->fileName;
                Utils::saveFile($outputStream, $outputPath);
                return $outputPath;
            } else {
                return false;
            }

        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }

    }

} 