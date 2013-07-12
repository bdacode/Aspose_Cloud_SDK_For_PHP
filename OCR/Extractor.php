<?php
namespace Aspose\Cloud\OCR;

use Aspose\Cloud\Common\Utils;
use Aspose\Cloud\Common\Product;
use Aspose\Cloud\Exception\AsposeCloudException as Exception;

class Extractor {

    //put your code here
    public function __construct() {
        
    }

    public function extractText() {
        $numOfArgs = func_get_args();
        switch (count($numOfArgs)) {
            case 1:
                $imageFileName = $numOfArgs[0];
                try {
                    $strURI = Product::$baseProductUri . '/ocr/' . $imageFileName . '/recognize?useDefaultDictionaries=true';
                    $signedURI = Utils::sign($strURI);
                    $responseStream = Utils::processCommand($signedURI, 'GET', '', '');
                    $json = json_decode($responseStream);
                    return $json;
                } catch (Exception $e) {
                    throw new Exception($e->getMessage());
                    return null;
                }
                break;
            case 2:
                $imageFileName = $numOfArgs[0];
                $folder = $numOfArgs[1];
                try {
                    if ($folder === '' || $folder === null) {
                        $strURI = Product::$baseProductUri . '/ocr/' . $imageFileName . '/recognize';
                    } else {
                        $strURI = Product::$baseProductUri . '/ocr/' . $imageFileName . '/recognize?folder=' . $folder;
                    }
                    $signedURI = Utils::sign($strURI);
                    $responseStream = Utils::processCommand($signedURI, 'GET', '', '');
                    $json = json_decode($responseStream);
                    return $json;
                } catch (Exception $e) {
                    throw new Exception($e->getMessage());
                    return null;
                }
                break;
            case 3:
                $stream = $numOfArgs[0];
                $language = $numOfArgs[1];
                $useDefaultDictionaries = $numOfArgs[2];
                try {
                    $strURI = Product::$baseProductUri . '/ocr/recognize?language=' . $language . '/recognize' . '&useDefaultDictionaries=';
                    $strURI .= ($useDefaultDictionaries) ? 'true' : 'false';
                    $signedURI = Utils::sign($strURI);
                    $responseStream = Utils::processCommand($signedURI, 'POST', '', $stream);
                    $json = json_decode($responseStream);
                    return $json;
                } catch (Exception $e) {
                    throw new Exception($e->getMessage());
                    return null;
                }
                break;
            case 4:
                $imageFileName = $numOfArgs[0];
                $folder = $numOfArgs[1];
                $language = $numOfArgs[2];
                $useDefaultDictionaries = $numOfArgs[3];

                try {
                    if ($folder === '' || $folder === null) {

                        $strURI = Product::$baseProductUri . '/ocr/' . $imageFileName . '/recognize?language=' . $language . '&useDefaultDictionaries=';
                        $strURI.= ($useDefaultDictionaries) ? 'true' : 'false';
                    } else {

                        $strURI = Product::$baseProductUri . '/ocr/' . $imageFileName . '/recognize?language=' . $language . '&useDefaultDictionaries=';
                        $strURI .= ($useDefaultDictionaries) ? 'true' : 'false';
                        $strURI .= '&folder=' . $folder;
                    }
                    $signedURI = Utils::sign($strURI);
                    $responseStream = Utils::processCommand($signedURI, 'GET', '', '');
                    $json = json_decode($responseStream);
                    return $json;
                } catch (Exception $e) {
                    throw new Exception($e->getMessage());
                    return null;
                }
                break;
            case 8:
                $imageFileName = $numOfArgs[0];
                $language = $numOfArgs[1];
                $useDefaultDictionaries = $numOfArgs[2];
                $x = $numOfArgs[3];
                $y = $numOfArgs[4];
                $height = $numOfArgs[5];
                $width = $numOfArgs[6];
                $folder = $numOfArgs[7];
                try {
                    $strURI = Product::$baseProductUri;
                    $strURI .= '/ocr/';
                    $strURI .= $imageFileName;
                    $strURI .= '/recognize?language=';
                    $strURI .= $language;
                    $strURI .= (($x >= 0 && $y >= 0 && $width > 0 && $height > 0) ? '&rectX=' . $x . '&rectY=' . $y . '&rectWidth=' . $width . '&rectHeight=' . $height : '');
                    $strURI .= '&useDefaultDictionaries=';
                    $strURI .= (($useDefaultDictionaries) ? 'true' : 'false');
                    $strURI .= (($folder === '') ? '' : '&folder=' + $folder);
                    $signedURI = Utils::sign($strURI);
                    $responseStream = Utils::processCommand($signedURI, 'GET', '', '');
                    $json = json_decode($responseStream);
                    return $json;
                } catch (Exception $e) {
                    throw new Exception($e->getMessage());
                    return null;
                }
                break;
            default :
                echo 'Wrong numbers of arguments';
                break;
        }
    }

    public function extractTextFromLocalFile($localFile, $language, $useDefaultDictionaries) {
        try {
            $strURI = Product::$baseProductUri . '/ocr/recognize?language=' . $language . '&useDefaultDictionaries=';
            $strURI .= ($useDefaultDictionaries) ? 'true' : 'false';
            $signedURI = Utils::sign($strURI);
            $stream = file_get_contents($localFile);
            $responseStream = Utils::processCommand($signedURI, 'POST', 'json', $stream);
            $json = json_decode($responseStream);
            return $json;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
            return null;
        }
    }

    public function extractTextFromUrl($url, $language, $useDefaultDictionaries) {
        try {
            $strURI = Product::$baseProductUri . '/ocr/recognize?url=' . $url . '&language=' . $language . '&useDefaultDictionaries=' . $useDefaultDictionaries;
            $signedURI = Utils::sign($strURI);
            $responseStream = Utils::processCommand($signedURI);
            $json = json_decode($responseStream);
            return $json;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

}
