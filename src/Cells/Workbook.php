<?php
/*
 * This class contains features to work with charts
 */
namespace Aspose\Cloud\Cells;

use Aspose\Cloud\Common\Utils;
use Aspose\Cloud\Common\Product;
use Aspose\Cloud\Exception\AsposeCloudException as Exception;

class Workbook {

    public $fileName = '';

    public function __construct($fileName) {
        $this->fileName = $fileName;
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
            $strURI = Product::$baseProductUri . '/cells/' . $this->fileName . '/documentProperties';
            //sign URI
            $signedURI = Utils::sign($strURI);
            $responseStream = Utils::processCommand($signedURI, 'GET', '', '');
            $json = json_decode($responseStream);
            if ($json->Code == 200)
                return $json->DocumentProperties->DocumentPropertyList;
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
            $strURI = Product::$baseProductUri . '/cells/' . $this->fileName . '/documentProperties/' . $propertyName;
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
            $strURI = Product::$baseProductUri . '/cells/' . $this->fileName . '/documentProperties/' . $propertyName;
            $put_data_arr['Value'] = $propertyValue;
            $put_data = json_encode($put_data_arr);
            //sign URI
            $signedURI = Utils::sign($strURI);
            $responseStream = Utils::processCommand($signedURI, 'PUT', 'json', $put_data);
            $json = json_decode($responseStream);
            if ($json->Code == 201) {
                return $json->DocumentProperty;
            } else {
                return false;
            }
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    /*
     * Remove All Document's properties
     */
    public function removeAllProperties() {
        try {
            //check whether files are set or not
            if ($this->fileName == '')
                throw new Exception('Base file not specified');
            //build URI to merge Docs
            $strURI = Product::$baseProductUri . '/cells/' . $this->fileName . '/documentProperties';
            //sign URI
            $signedURI = Utils::sign($strURI);
            $responseStream = Utils::processCommand($signedURI, 'DELETE');
            $json = json_decode($responseStream);
            if (is_object($json)) {
                if ($json->Code == 200)
                    return true;
                else
                    return false;
            }
            return true;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    /*
     * Delete a document property
      @param string $propertyName
     */
    public function removeProperty($propertyName) {
        try {
            //check whether files are set or not
            if ($this->fileName == '')
                throw new Exception('Base file not specified');
            if ($propertyName == '')
                throw new Exception('Property Name not specified');
            //build URI to merge Docs
            $strURI = Product::$baseProductUri . '/cells/' . $this->fileName . '/documentProperties/' . $propertyName;
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
     * Create Empty Workbook
     */
    public function createEmptyWorkbook() {
        try {
            //build URI to merge Docs
            $strURI = Product::$baseProductUri . '/cells/' . $this->fileName;
            //sign URI
            $signedURI = Utils::sign($strURI);
            $responseStream = Utils::processCommand($signedURI, 'PUT');
            $json = json_decode($responseStream);
            return $json;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    /*
     * Create Empty Workbook
     * @param string $templateFileName
     */
    public function createWorkbookFromTemplate($templateFileName) {
        try {
            if ($templateFileName == '')
                throw new Exception('Template file not specified');
            //build URI to merge Docs
            $strURI = Product::$baseProductUri . '/cells/' . $this->fileName . '?templatefile=' . $templateFileName;
            //sign URI
            $signedURI = Utils::sign($strURI);
            $responseStream = Utils::processCommand($signedURI, 'PUT');
            $json = json_decode($responseStream);
            return $json;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    /*
     * Create Empty Workbook
     * @param string $templateFileName
     * @param string $dataFile	
     */
    public function createWorkbookFromSmartMarkerTemplate($templateFileName, $dataFile) {
        try {
            if ($templateFileName == '')
                throw new Exception('Template file not specified');
            if ($dataFile == '')
                throw new Exception('Data file not specified');
            //build URI to merge Docs
            $strURI = Product::$baseProductUri . '/cells/' . $this->fileName . '?templatefile=' . $templateFileName . '&dataFile=' . $dataFile;
            //sign URI
            $signedURI = Utils::sign($strURI);
            $responseStream = Utils::processCommand($signedURI, 'PUT');
            $json = json_decode($responseStream);
            return $json;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    /*
     * Process Smartmaker Datafile
     * @param string $dataFile	
     */
    public function processSmartMarker($dataFile) {
        try {
            if ($dataFile == '')
                throw new Exception('Data file not specified');
            //build URI to merge Docs
            $strURI = Product::$baseProductUri . '/cells/' . $this->fileName . '/smartmarker?xmlFile=' . $dataFile;
            //sign URI
            $signedURI = Utils::sign($strURI);
            $responseStream = Utils::processCommand($signedURI, 'POST');
            $json = json_decode($responseStream);
            return $json;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    /*
     * Get Worksheets Count in Workbook
     */
    public function getWorksheetsCount() {
        try {
            if ($this->fileName == '')
                throw new Exception('Base file not specified');
            //build URI to merge Docs
            $strURI = Product::$baseProductUri . '/cells/' . $this->fileName . '/worksheets';
            //sign URI
            $signedURI = Utils::sign($strURI);
            $responseStream = Utils::processCommand($signedURI, 'GET', '', '');
            $json = json_decode($responseStream);
            return count($json->Worksheets->WorksheetList);
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    /*
     * Get Names Count in Workbook	
     */
    public function getNamesCount() {
        try {
            if ($this->fileName == '')
                throw new Exception('Base file not specified');
            //build URI to merge Docs
            $strURI = Product::$baseProductUri . '/cells/' . $this->fileName . '/names';
            //sign URI
            $signedURI = Utils::sign($strURI);
            $responseStream = Utils::processCommand($signedURI, 'GET', '', '');
            $json = json_decode($responseStream);
            return $json->Names->Count;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    /*
     * Get Default Style
     */
    public function getDefaultStyle() {
        try {
            if ($this->fileName == '')
                throw new Exception('Base file not specified');
            //build URI to merge Docs
            $strURI = Product::$baseProductUri . '/cells/' . $this->fileName . '/defaultStyle';
            //sign URI
            $signedURI = Utils::sign($strURI);
            $responseStream = Utils::processCommand($signedURI, 'GET', '');
            $json = json_decode($responseStream);
            return $json->Style;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    public function encryptWorkbook($encryptionType = 'XOR', $password = '', $keyLength = '') {
        try {
            if ($this->fileName == '')
                throw new Exception('Base file not specified');
            //Build JSON to post
            $fieldsArray['EncriptionType'] = $encryptionType;
            $fieldsArray['KeyLength'] = $keyLength;
            $fieldsArray['Password'] = $password;
            $json = json_encode($fieldsArray);
            $strURI = Product::$baseProductUri . '/cells/' . $this->fileName . '/encryption';
            $signedURI = Utils::sign($strURI);
            $responseStream = Utils::processCommand($signedURI, 'POST', 'json', $json);
            $json_response = json_decode($responseStream);
            if ($json_response->Code == 200)
                return true;
            else
                return false;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    public function protectWorkbook($password, $protectionType = 'all') {
        try {
            if ($this->fileName == '')
                throw new Exception('Base file not specified');
            //Build JSON to post
            $fieldsArray['ProtectionType'] = $protectionType;
            $fieldsArray['Password'] = $password;
            $json = json_encode($fieldsArray);
            $strURI = Product::$baseProductUri . '/cells/' . $this->fileName . '/protection';
            $signedURI = Utils::sign($strURI);
            $responseStream = Utils::processCommand($signedURI, 'POST', 'json', $json);
            $json_response = json_decode($responseStream);
            if ($json_response->Code == 200)
                return true;
            else
                return false;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    public function unprotectWorkbook($password) {
        try {
            if ($this->fileName == '')
                throw new Exception('Base file not specified');
            //Build JSON to post			
            $fieldsArray['Password'] = $password;
            $json = json_encode($fieldsArray);
            $strURI = Product::$baseProductUri . '/cells/' . $this->fileName . '/protection';
            $signedURI = Utils::sign($strURI);
            $responseStream = Utils::processCommand($signedURI, 'DELETE', 'json', $json);
            $json_response = json_decode($responseStream);
            if ($json_response->Code == 200)
                return true;
            else
                return false;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    public function setModifyPassword($password) {
        try {
            if ($this->fileName == '')
                throw new Exception('Base file not specified');
            //Build JSON to post
            $fieldsArray['Password'] = $password;
            $json = json_encode($fieldsArray);
            $strURI = Product::$baseProductUri . '/cells/' . $this->fileName . '/writeProtection';
            $signedURI = Utils::sign($strURI);
            $responseStream = Utils::processCommand($signedURI, 'PUT', 'json', $json);
            $json_response = json_decode($responseStream);
            if ($json_response->Status == 'OK')
                return true;
            else
                return false;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    public function clearModifyPassword($password) {
        try {
            if ($this->fileName == '')
                throw new Exception('Base file not specified');
            //Build JSON to post
            $fieldsArray['Password'] = $password;
            $json = json_encode($fieldsArray);
            $strURI = Product::$baseProductUri . '/cells/' . $this->fileName . '/writeProtection';
            $signedURI = Utils::sign($strURI);
            $responseStream = Utils::processCommand($signedURI, 'DELETE', 'json', $json);
            $json_response = json_decode($responseStream);
            if ($json_response->Status == 'OK')
                return true;
            else
                return false;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    public function decryptWorkbook($password) {
        try {
            if ($this->fileName == '')
                throw new Exception('Base file not specified');
            //Build JSON to post
            $fieldsArray['Password'] = $password;
            $json = json_encode($fieldsArray);
            $strURI = Product::$baseProductUri . '/cells/' . $this->fileName . '/encryption';
            $signedURI = Utils::sign($strURI);
            $responseStream = Utils::processCommand($signedURI, 'DELETE', 'json', $json);
            $json_response = json_decode($responseStream);
            if ($json_response->Code == 200)
                return true;
            else
                return false;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    public function addWorksheet($worksheetName) {
        try {
            if ($this->fileName == '')
                throw new Exception('Base file not specified');
            $strURI = Product::$baseProductUri . '/cells/' . $this->fileName . '/worksheets/' . $worksheetName;
            $signedURI = Utils::sign($strURI);
            $responseStream = Utils::processCommand($signedURI, 'PUT', '', '');
            $json_response = json_decode($responseStream);
            if ($json_response->Code == 201)
                return true;
            else
                return false;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    public function removeWorksheet($worksheetName) {
        try {
            if ($this->fileName == '')
                throw new Exception('Base file not specified');
            $strURI = Product::$baseProductUri . '/cells/' . $this->fileName . '/worksheets/' . $worksheetName;
            $signedURI = Utils::sign($strURI);
            $responseStream = Utils::processCommand($signedURI, 'DELETE', '', '');
            $json_response = json_decode($responseStream);
            if ($json_response->Code == 200)
                return true;
            else
                return false;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    public function mergeWorkbook($mergeFileName) {
        try {
            if ($this->fileName == '')
                throw new Exception('Base file not specified');
            $strURI = Product::$baseProductUri . '/cells/' . $this->fileName . '/merge?mergeWith=' . $mergeFileName;
            $signedURI = Utils::sign($strURI);
            $responseStream = Utils::processCommand($signedURI, 'POST', '', '');
            $json_response = json_decode($responseStream);
            if ($json_response->Code == 200)
                return true;
            else
                return false;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

}