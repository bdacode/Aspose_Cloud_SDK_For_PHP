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

class Field {
    /*
     * Inserts page number filed into the document.
     * @param string $fileName 
     * @param string $alignment
     * @param string $format 
     * @param string $isTop
     * @param string $setPageNumberOnFirstPage 
     */
    public function insertPageNumber($fileName, $alignment, $format, $isTop, $setPageNumberOnFirstPage) {
        try {
            //check whether files are set or not
            if ($fileName == '')
                throw new Exception('File not specified');

            //Build JSON to post
            $fieldsArray = array('Format' => $format, 'Alignment' => $alignment,
                'IsTop' => $isTop, 'SetPageNumberOnFirstPage' => $setPageNumberOnFirstPage);
            $json = json_encode($fieldsArray);

            //build URI to insert page number
            $strURI = Product::$baseProductUri . '/words/' . $fileName . '/insertPageNumbers';

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
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    /*
     * Gets all merge filed names from document.
     * @param string $fileName  
     */
    public function getMailMergeFieldNames($fileName) {
        try {
            //check whether file is set or not
            if ($fileName == '')
                throw new Exception('No file name specified');

            $strURI = Product::$baseProductUri . '/words/' . $fileName . '/mailMergeFieldNames';

            $signedURI = Utils::sign($strURI);

            $responseStream = Utils::processCommand($signedURI, 'GET', '', '');

            $json = json_decode($responseStream);
            
            return $json->FieldNames->Names;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

}