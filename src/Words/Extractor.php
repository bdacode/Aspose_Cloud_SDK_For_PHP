<?php
/*
 * Extract various types of information from the document
 */
namespace Aspose\Cloud\Words;

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
     * Gets Text items list from document
     */
    public function getText() {
        try {
            //check whether file is set or not
            if ($this->fileName == '')
                throw new Exception('No file name specified');

            $strURI = Product::$baseProductUri . '/words/' . $this->fileName . '/textItems';

            $signedURI = Utils::sign($strURI);

            $responseStream = Utils::processCommand($signedURI, 'GET', '', '');

            $json = json_decode($responseStream);

            return $json->TextItems->List;
            
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    /*
     * Get the OLE drawing object from document
     * @param int $index
     * @param string $OLEFormat
     */
    public function getoleData($index, $OLEFormat) {
        try {
            //check whether file is set or not
            if ($this->fileName == '')
                throw new Exception('No file name specified');

            $strURI = Product::$baseProductUri . '/words/' . $this->fileName . '/drawingObjects/' . $index . '/oleData';

            $signedURI = Utils::sign($strURI);

            $responseStream = Utils::processCommand($signedURI, 'GET', '', '');

            $v_output = Utils::validateOutput($responseStream);

            if ($v_output === '') {
                $outputPath =  AsposeApp::$outPutLocation . Utils::getFileName($this->fileName) . '_' . $index . '.' . $OLEFormat;
                Utils::saveFile($responseStream,$outputPath);
                return $outputPath;
            }
            else
                return $v_output;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    /*
     * Get the Image drawing object from document
     * @param int $index
     * @param string $renderformat
     */
    public function getimageData($index, $renderFormat) {
        try {
            //check whether file is set or not
            if ($this->fileName == '')
                throw new Exception('No file name specified');

            $strURI = Product::$baseProductUri . '/words/' . $this->fileName . '/drawingObjects/' . $index . '/ImageData';

            $signedURI = Utils::sign($strURI);

            $responseStream = Utils::processCommand($signedURI, 'GET', '', '');

            $v_output = Utils::validateOutput($responseStream);

            if ($v_output === '') {
                $outputPath = AsposeApp::$outPutLocation . Utils::getFileName($this->fileName) . '_' . $index . '.' . $renderFormat;
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
     * Convert drawing object to image
     * @param int $index
     * @param string $renderformat
     */
    public function convertDrawingObject($index, $renderFormat) {
        try {
            //check whether file is set or not
            if ($this->fileName == '')
                throw new Exception('No file name specified');

            $strURI = Product::$baseProductUri . '/words/' . $this->fileName . '/drawingObjects/' . $index . '?format=' . $renderFormat;

            $signedURI = Utils::sign($strURI);

            $responseStream = Utils::processCommand($signedURI, 'GET', '', '');

            $v_output = Utils::validateOutput($responseStream);

            if ($v_output === '') {
                $outputPath = AsposeApp::$outPutLocation . Utils::getFileName($this->fileName) . '_' . $index . '.' . $renderFormat;
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
     * Get the List of drawing object from document	
     */
    public function getDrawingObjectList() {
        try {
            //check whether file is set or not
            if ($this->fileName == '')
                throw new Exception('No file name specified');

            $strURI = Product::$baseProductUri . '/words/' . $this->fileName . '/drawingObjects';

            $signedURI = Utils::sign($strURI);

            $responseStream = Utils::processCommand($signedURI, 'GET', '', '');

            $json = json_decode($responseStream);

            if ($json->Code == 200)
                return $json->DrawingObjects->List;
            else
                return false;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    /*
     * Get the drawing object from document	
     * @param string $objectURI
     * @param string $outputPath
     */
    public function getDrawingObject($objectURI, $outputPath) {
        try {
            //check whether file is set or not
            if ($this->fileName == '')
                throw new Exception('No file name specified');

            if ($outputPath == '')
                throw new Exception('Output path not specified');

            if ($objectURI == '')
                throw new Exception('Object URI not specified');

            $url_arr = explode('/', $objectURI);
            $objectIndex = end($url_arr);

            $strURI = $objectURI;

            $signedURI = Utils::sign($strURI);

            $responseStream = Utils::processCommand($signedURI, 'GET', '', '');

            $json = json_decode($responseStream);

            if ($json->Code == 200) {
                if ($json->DrawingObject->ImageDataLink != '') {
                    $strURI = $strURI . '/imageData';
                    $outputPath = $outputPath . '\\DrawingObject_' . $objectIndex . '.jpeg';
                } else if ($json->DrawingObject->OleDataLink != '') {
                    $strURI = $strURI . '/oleData';
                    $outputPath = $outputPath . '\\DrawingObject_' . $objectIndex . '.xlsx';
                } else {
                    $strURI = $strURI . '?format=jpeg';
                    $outputPath = $outputPath . '\\DrawingObject_' . $objectIndex . '.jpeg';
                }

                $signedURI = Utils::sign($strURI);

                $responseStream = Utils::processCommand($signedURI, 'GET', '', '');

                $v_output = Utils::validateOutput($responseStream);

                if ($v_output === '') {
                    Utils::saveFile($responseStream, $outputPath);
                    return $outputPath;
                }
                else
                    return $v_output;
            }
            else {
                return false;
            }
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    /*
     * Get the List of drawing object from document	
     * @param string outputPath
     */
    public function getDrawingObjects($outputPath) {
        try {
            //check whether file is set or not
            if ($this->fileName == '')
                throw new Exception('No file name specified');

            if ($outputPath == '')
                throw new Exception('Output path not specified');

            $strURI = Product::$baseProductUri . '/words/' . $this->fileName . '/drawingObjects';

            $signedURI = Utils::sign($strURI);

            $responseStream = Utils::processCommand($signedURI, 'GET', '', '');

            $json = json_decode($responseStream);

            if ($json->Code == 200) {
                foreach ($json->DrawingObjects->List as $object) {
                    $this->GetDrawingObject($object->link->Href, $outputPath);
                }
            }
            else
                return false;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

}