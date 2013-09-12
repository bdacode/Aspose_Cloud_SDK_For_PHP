<?php
/*
 * Deals with Word document level aspects
 */
namespace Aspose\Cloud\Words;

use Aspose\Cloud\Common\AsposeApp;
use Aspose\Cloud\Common\Utils;
use Aspose\Cloud\Common\Product;
use Aspose\Cloud\Storage\Folder;
use Aspose\Cloud\Exception\AsposeCloudException as Exception;

class Document {

    public $fileName = '';

    public function __construct($fileName) {
        $this->fileName = $fileName;
    }

    /*
     * Appends a list of documents to this one.
     * @param string $appendDocs (List of documents to append)
     * @param string $importFormatModes
     * @param string $sourceFolder (name of the folder where documents are present)
     */
    public function appendDocument($appendDocs, $importFormatModes, $sourceFolder) {
        try {
            //check whether files are set or not
            if ($this->fileName == '')
                throw new Exception('Base file not specified');
            //check whether required information is complete
            if (count($appendDocs) != count($importFormatModes))
                throw new Exception('Please specify complete documents and import format modes');

            $post_array = array();
            $i = 0;
            foreach ($appendDocs as $doc) {
                $post_array[] = array("Href" => (($sourceFolder != "" ) ? $sourceFolder . "\\" . $doc : $doc), "ImportFormatMode" => $importFormatModes[$i]);
                $i++;
            }
            $data = array("DocumentEntries" => $post_array);
            $json = json_encode($data);

            //build URI to merge Docs
            $strURI = Product::$baseProductUri . '/words/' . $this->fileName . '/appendDocument';

            //sign URI
            $signedURI = Utils::sign($strURI);

            $responseStream = Utils::processCommand($signedURI, 'POST', 'json', $json);

            $v_output = Utils::validateOutput($responseStream);

            if ($v_output === '') {
                //Save merged docs on server
                $folder = new Folder();
                $outputStream = $folder->GetFile($sourceFolder . (($sourceFolder == '') ? '' : '/') . $this->fileName);
                $outputPath = AsposeApp::$outPutLocation . $this->fileName;
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
     * Get Resource Properties information like document source format, IsEncrypted, IsSigned and document properties
     */
    public function getDocumentInfo() {
        try {
            //check whether files are set or not
            if ($this->fileName == '')
                throw new Exception('Base file not specified');

            //build URI to merge Docs
            $strURI = Product::$baseProductUri . '/words/' . $this->fileName;

            //sign URI
            $signedURI = Utils::sign($strURI);

            $responseStream = Utils::processCommand($signedURI, 'GET', '', '');

            $json = json_decode($responseStream);

            if ($json->Code == 200)
                return $json->Document;
            else
                return false;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    /*
     * Get Resource Properties information like document source format, IsEncrypted, IsSigned and document properties
      @param string $propertyName
     */
    public function getProperty($propertyName) {
        try {
            //check whether files are set or not
            if ($this->fileName == '')
                throw new Exception('Base file not specified');

            if ($propertyName == '')
                throw new Exception('Property Name not specified');

            //build URI to merge Docs
            $strURI = Product::$baseProductUri . '/words/' . $this->fileName . '/documentProperties/' . $propertyName;

            //sign URI
            $signedURI = Utils::sign($strURI);

            $responseStream = Utils::processCommand($signedURI, 'GET', '', '');

            $json = json_decode($responseStream);


            if ($json->Code == 200)
                return $json->DocumentProperty;
            else
                return false;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    /*
     * Set document property
      @param string $propertyName
      @param string $propertyValue
     */
    public function setProperty($propertyName, $propertyValue) {
        try {
            //check whether files are set or not
            if ($this->fileName == '')
                throw new Exception('Base file not specified');

            if ($propertyName == '')
                throw new Exception('Property Name not specified');

            if ($propertyValue == '')
                throw new Exception('Property Value not specified');

            //build URI to merge Docs
            $strURI = Product::$baseProductUri . '/words/' . $this->fileName . '/documentProperties/' . $propertyName;

            $put_data_arr['Value'] = $propertyValue;

            $put_data = json_encode($put_data_arr);

            //sign URI
            $signedURI = Utils::sign($strURI);

            $responseStream = Utils::processCommand($signedURI, 'PUT', 'json', $put_data);

            $json = json_decode($responseStream);

            if ($json->Code == 200)
                return $json->DocumentProperty;
            else
                return false;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    public function protectDocument($password, $protectionType = 'AllowOnlyComments') {
        try {
            if ($this->fileName == '') {
                throw new Exception('Base file not specified');
            }
            if ($password == '') {
                throw new Exception('Please Specify A Password');
            }
            $fieldsArray = array('Password' => $password, 'ProtectionType' => $protectionType);
            $json = json_encode($fieldsArray);
            $strURI = Product::$baseProductUri . '/words/' . $this->fileName . '/protection';
            $signedURI = Utils::sign($strURI);
            $responseStream = Utils::processCommand($signedURI, 'PUT', 'json', $json);
            $v_output = Utils::validateOutput($responseStream);
            if ($v_output === '') {
                $strURI = Product::$baseProductUri . '/storage/file/' . $this->fileName;
                $signedURI = Utils::sign($strURI);
                $responseStream = Utils::processCommand($signedURI, 'GET', '', '');
                $outputFile = AsposeApp::$outPutLocation . $this->fileName;
                Utils::saveFile($responseStream, $outputFile);
                return $outputFile;
            }
            else
                return $v_output;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    public function unprotectDocument($password, $protectionType = 'AllowOnlyComments') {
        try {
            if ($this->fileName == '') {
                throw new Exception('Base file not specified');
            }
            if ($password == '') {
                throw new Exception('Please Specify A Password');
            }
            $fieldsArray = array('Password' => $password, 'ProtectionType' => $protectionType);
            $json = json_encode($fieldsArray);
            $strURI = Product::$baseProductUri . '/words/' . $this->fileName . '/protection';
            $signedURI = Utils::sign($strURI);
            $responseStream = Utils::processCommand($signedURI, 'DELETE', 'json', $json);
            $v_output = Utils::validateOutput($responseStream);
            if ($v_output === '') {
                $strURI = Product::$baseProductUri . '/storage/file/' . $this->fileName;
                $signedURI = Utils::sign($strURI);
                $responseStream = Utils::processCommand($signedURI, 'GET', '', '');
                $outputFile = AsposeApp::$OutPutLocation . $this->FileName;
                Utils::saveFile($responseStream, $outputFile);
                return $outputFile;
            }
            else
                return $v_output;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    public function updateProtection($oldPassword, $newPassword, $protectionType = 'AllowOnlyComments') {
        try {
            if ($this->fileName == '') {
                throw new Exception('Base file not specified');
            }
            if ($oldPassword == '') {
                throw new Exception('Please Specify Old Password');
            }
            if ($newPassword == '') {
                throw new Exception('Please Specify New Password');
            }
            $fieldsArray = array('Password' => $oldPassword, 'NewPassword' => $newPassword, 'ProtectionType' => $protectionType);
            $json = json_encode($fieldsArray);
            $strURI = Product::$baseProductUri . '/words/' . $this->fileName . '/protection';
            $signedURI = Utils::sign($strURI);
            $responseStream = Utils::processCommand($signedURI, 'POST', 'json', $json);
            $v_output = Utils::validateOutput($responseStream);
            if ($v_output === '') {
                $strURI = Product::$baseProductUri . '/storage/file/' . $this->fileName;
                $signedURI = Utils::sign($strURI);
                $responseStream = Utils::processCommand($signedURI, 'GET', '', '');
                $outputFile = AsposeApp::$outPutLocation . $this->fileName;
                Utils::saveFile($responseStream, $outputFile);
                return $outputFile;
            }
            else
                return $v_output;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    /*
     * Delete a document property
      @param string $propertyName
     */
    public function deleteProperty($propertyName) {
        try {
            //check whether files are set or not
            if ($this->fileName == '')
                throw new Exception('Base file not specified');

            if ($propertyName == '')
                throw new Exception('Property Name not specified');

            //build URI to merge Docs
            $strURI = Product::$baseProductUri . '/words/' . $this->fileName . '/documentProperties/' . $propertyName;

            //sign URI
            $signedURI = Utils::sign($strURI);
            
            $responseStream = Utils::processCommand($signedURI, 'DELETE', '', '');

            $json = json_decode($responseStream);

            if ($json->Code == 200)
                return true;
            else
                return false;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    /*
     * Get Document's properties
     */
    public function getProperties() {
        try {
            //check whether files are set or not
            if ($this->fileName == '')
                throw new Exception('Base file not specified');


            //build URI to merge Docs
            $strURI = Product::$baseProductUri . '/words/' . $this->fileName . '/documentProperties';

            //sign URI
            $signedURI = Utils::sign($strURI);

            $responseStream = Utils::processCommand($signedURI, 'GET', '', '');

            $json = json_decode($responseStream);


            if ($json->Code == 200)
                return $json->DocumentProperties->List;
            else
                return false;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    /*
     * Convert Document to different file format without using storage
     * $param string $inputPath
     * @param string $outputPath
     * @param string $outputFormat
     */
    public function convertLocalFile($inputPath = '', $outputPath = '', $outputFormat = '') {
        try {
            //check whether file is set or not
            if ($inputPath == '')
                throw new Exception('No file name specified');

            if ($outputFormat == '')
                throw new Exception('output format not specified');


            $strURI = Product::$baseProductUri . '/words/convert?format=' . $outputFormat;

            if (!file_exists($inputPath)) {
                throw new Exception('input file doesnt exist.');
            }


            $signedURI = Utils::sign($strURI);

            $responseStream = Utils::uploadFileBinary($signedURI, $inputPath, 'xml');

            $v_output = Utils::validateOutput($responseStream);

            if ($v_output === '') {

                $save_format = $outputFormat;

                if ($outputPath == '') {
                    $outputPath = Utils::getFileName($inputPath) . '.' . $save_format;
                }
                $output =  AsposeApp::$outPutLocation . $outputPath;
                Utils::saveFile($responseStream,$output);
                return true;
            }
            else
                return $v_output;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

}