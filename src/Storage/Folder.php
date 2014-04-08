<?php
/*
 *  Main class that provides methods to perform all the transactions on the storage of a Aspose Application.
 */
namespace Aspose\Cloud\Storage;

use Aspose\Cloud\Common\Utils;
use Aspose\Cloud\Common\Product;
use Aspose\Cloud\Exception\AsposeCloudException as Exception;

class Folder {

    public $strURIFolder = '';
    public $strURIFile = '';
    public $strURIExist = '';
    public $strURIDisc = '';

    public function __construct() {
        $this->strURIFolder = Product::$baseProductUri . '/storage/folder/';
        $this->strURIFile = Product::$baseProductUri . '/storage/file/';
        $this->strURIExist = Product::$baseProductUri . '/storage/exist/';
        $this->strURIDisc = Product::$baseProductUri . '/storage/disc';
    }

    /*
     * Uploads a file from your local machine to specified folder / subfolder on Aspose storage.
     * 
     * @param string $strFile
     * @param string $strFolder
     */
    public function uploadFile($strFile, $strFolder, $storageName = '') {
        $strRemoteFileName = basename($strFile);
        $strURIRequest = $this->strURIFile;
        if ($strFolder == '')
            $strURIRequest .= $strRemoteFileName;
        else
            $strURIRequest .= $strFolder . '/' . $strRemoteFileName;
        if ($storageName != '') {
            $strURIRequest .= '?storage=' . $storageName;
        }
        $signedURI = Utils::sign($strURIRequest);

        Utils::uploadFileBinary($signedURI, $strFile);
    }

    /*
     * Checks if a file exists
     *
     * @param string $fileName
     */
    public function fileExists($fileName, $storageName = '') {
        //check whether file is set or not
        if ($fileName == '')
            throw new Exception('No file name specified');

        //build URI
        $strURI = $this->strURIExist . $fileName;
        if ($storageName != '') {
            $strURI .= '?storage=' . $storageName;
        }
        //sign URI
        $signedURI = Utils::sign($strURI);

        $responseStream = json_decode(Utils::processCommand($signedURI, 'GET', '', ''));
        if (!$responseStream->FileExist->IsExist) {
            return FALSE;
        }
        return TRUE;
    }

    /*
     * Deletes a file from remote storage
     *
     * @param string $fileName
     */
    public function deleteFile($fileName, $storageName = '') {
        //check whether file is set or not
        //build URI
        if ($fileName == '')
            throw new Exception('No file name specified');
        $strURI = $this->strURIFile . $fileName;
        if ($storageName != '') {
            $strURI .= '?storage=' . $storageName;
        }
        //sign URI
        $signedURI = Utils::sign($strURI);

        $responseStream = json_decode(Utils::processCommand($signedURI, 'DELETE', '', ''));
        if ($responseStream->Code != 200) {
            return FALSE;
        }
        return TRUE;
    }

    /*
     * Creates a new folder  under the specified folder on Aspose storage. If no path specified, creates a folder under the root folder.
     * 
     * @param string $strFolder
     */
    public function createFolder($strFolder, $storageName = '') {
        //build URI
        $strURIRequest = $this->strURIFolder . $strFolder;
        if ($storageName != '') {
            $strURIRequest .= '?storage=' . $storageName;
        }
        //sign URI
        $signedURI = Utils::sign($strURIRequest);

        $responseStream = json_decode(Utils::processCommand($signedURI, 'PUT', '', ''));

        if ($responseStream->Code != 200) {
            return FALSE;
        }
        return TRUE;
    }

    /*
     * Deletes a folder from remote storage
     *
     * @param string $folderName
     */
    public function deleteFolder($folderName) {
        //check whether folder is set or not
        if ($folderName == '')
            throw new Exception('No folder name specified');

        //build URI
        $strURI = $this->strURIFolder . $folderName;

        //sign URI
        $signedURI = Utils::sign($strURI);

        $responseStream = json_decode(Utils::processCommand($signedURI, 'DELETE', '', ''));
        if ($responseStream->Code != 200) {
            return FALSE;
        }
        return TRUE;
    }

    /*
     * Provides the total / free disc size in bytes for your app
     */
    public function getDiscUsage($storageName = '') {
            //build URI
            $strURI = $this->strURIDisc;
            if ($storageName != '') {
                $strURI .= '?storage=' . $storageName;
            }
            //sign URI
            $signedURI = Utils::sign($strURI);

            $responseStream = json_decode(Utils::processCommand($signedURI, 'GET', '', ''));

            return $responseStream->DiscUsage;
    }

    /*
     * Get file from Aspose server
     *
     * @param string $fileName
     */
    public function getFile($fileName, $storageName = '') {
        //check whether file is set or not
        if ($fileName == '')
            throw new Exception('No file name specified');

        //build URI
        $strURI = $this->strURIFile . $fileName;
        if ($storageName != '') {
            $strURI .= '?storage=' . $storageName;
        }
        //sign URI
        $signedURI = Utils::sign($strURI);

        $responseStream = Utils::processCommand($signedURI, 'GET', '', '');

        return $responseStream;
    }

    /*
     * Retrives the list of files and folders under the specified folder. Use empty string to specify root folder.
     *
     * @param string $strFolder
     */
    public function getFilesList($strFolder, $storageName = '') {
        //build URI
        $strURI = $this->strURIFolder;
        //check whether file is set or not
        if (!$strFolder == '') {
            $strURI .= $strFolder;
        }
        if ($storageName != '') {
            $strURI .='?storage=' . $storageName;
        }

        //sign URI
        $signedURI = Utils::sign($strURI);

        $responseStream = Utils::processCommand($signedURI, 'GET', '', '');

        $json = json_decode($responseStream);

        return $json->Files;
    }

}