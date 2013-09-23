<?php
/*
 * Deals with PDF document level aspects
 */
namespace Aspose\Cloud\Pdf;

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
     * Gets the page count of the specified PDF document
     */
    public function getPageCount() {
        //build URI
        $strURI = Product::$baseProductUri . '/pdf/' . $this->fileName . '/pages';

        //sign URI
        $signedURI = Utils::sign($strURI);

        //get response stream
        $responseStream = Utils::ProcessCommand($signedURI, 'GET', '');

        $json = json_decode($responseStream);

        return count($json->Pages->List);
    }

    /*
     * Merges two PDF documents
     * @param string $basePdf (name of the base/first PDF file)
     * @param string $newPdf (name of the second PDF file to merge with base PDF file)
     * @param string $startPage (page number to start merging second PDF: enter 0 to merge complete document)
     * @param string $endPage (page number to end merging second PDF: enter 0 to merge complete document)
     * @param string $sourceFolder (name of the folder where base/first and second input PDFs are present)
     */
    public function appendDocument($basePdf, $newPdf, $startPage = 0, $endPage = 0, $sourceFolder = '') {
        try {
            //check whether files are set or not
            if ($basePdf == '')
                throw new Exception('Base file not specified');
            if ($newPdf == '')
                throw new Exception('File to merge is not specified');

            //build URI to merge PDFs
            if ($sourceFolder == '')
                $strURI = Product::$baseProductUri . '/pdf/' . $basePdf .
                        '/appendDocument?appendFile=' . $newPdf . ($startPage > 0 ? '&startPage=' . $startPage : '' ) .
                        ($endPage > 0 ? '&endPage=' . $endPage : '' );
            else
                $strURI = Product::$baseProductUri . '/pdf/' . $basePdf .
                        '/appendDocument?appendFile=' . $sourceFolder . '/' . $newPdf .
                        ($startPage > 0 ? '&startPage=' . $startPage : '' ) .
                        ($endPage > 0 ? '&endPage=' . $endPage : '' ) .
                        '&folder=' . $sourceFolder;

            //sign URI
            $signedURI = Utils::sign($strURI);

            $responseStream = Utils::processCommand($signedURI, 'POST', '', '');

            $json = json_decode($responseStream);
           
            if ($json->Code == 200) {
                $folder = new Folder();
                $path = "";
                if($sourceFolder == ""){
                    $path = $basePdf;
                }else{
                    $path = $sourceFolder . '/' . $basePdf;
                }
                $outputStream = $folder->GetFile($path);
                $outputPath = AsposeApp::$outPutLocation . $basePdf;
                Utils::saveFile($outputStream, $outputPath);
            } else {
                return false;
            }
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    /*
     * Merges tow or more PDF documents
     * @param array $sourceFiles (list of PDF files to be merged)
     */
    public function mergeDocuments(array $sourceFiles = array()) {
        try {
            $mergedFileName = $this->fileName;
            //check whether files are set or not
            if ($mergedFileName == '')
                throw new Exception('Output file not specified');
            if (empty($sourceFiles))
                throw new Exception('File to merge are not specified');
            if (count($sourceFiles) < 2)
                throw new Exception('Two or more files are requred to merge');


            //Build JSON to post
            $documentsList = array('List' => $sourceFiles);
            $json = json_encode($documentsList);

            $strURI = Product::$baseProductUri . '/pdf/' . $mergedFileName . '/merge';

            //sign URI
            $signedURI = Utils::sign($strURI);

            $responseStream = json_decode(Utils::processCommand($signedURI, 'PUT', 'json', $json));

            if ($responseStream->Code == 200)
                return true;
            else
                return false;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    /*
     * Creates a PDF from HTML
     * @param string $pdfFileName (name of the PDF file to create)
     * @param string $htmlFileName (name of the HTML template file)
     */
    public function createFromHtml($pdfFileName, $htmlFileName) {
        try {
            //check whether files are set or not
            if ($pdfFileName == '')
                throw new Exception('PDF file name not specified');
            if ($htmlFileName == '')
                throw new Exception('HTML template file name not specified');

            //build URI to create PDF
            $strURI = Product::$baseProductUri . '/pdf/' . $pdfFileName .
                    '?templateFile=' . $htmlFileName . '&templateType=html';

            //sign URI
            $signedURI = Utils::sign($strURI);

            $responseStream = Utils::processCommand($signedURI, 'PUT', '', '');

            $v_output = Utils::validateOutput($responseStream);

            if ($v_output === '') {
                //Save PDF file on server
                $folder = new Folder();
                $outputStream = $folder->GetFile($pdfFileName);
                $outputPath = AsposeApp::$outPutLocation . $pdfFileName;
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
     * Creates a PDF from XML
     * @param string $pdfFileName (name of the PDF file to create)
     * @param string $xsltFileName (name of the XSLT template file)
     * @param string $xmlFileName (name of the XML file)
     */
    public function createFromXml($pdfFileName, $xsltFileName, $xmlFileName) {
        try {
            //check whether files are set or not
            if ($pdfFileName == '')
                throw new Exception('PDF file name not specified');
            if ($xsltFileName == '')
                throw new Exception('XSLT file name not specified');
            if ($xmlFileName == '')
                throw new Exception('XML file name not specified');

            //build URI to create PDF
            $strURI = Product::$baseProductUri . '/pdf/' . $pdfFileName . '?templateFile=' .
                    $xsltFileName . '&dataFile=' . $xmlFileName . '&templateType=xml';

            //sign URI
            $signedURI = Utils::sign($strURI);

            $responseStream = Utils::processCommand($signedURI, 'PUT', '', '');

            $v_output = Utils::validateOutput($responseStream);

            if ($v_output === '') {
                //Save PDF file on server
                $folder = new Folder();
                $outputStream = $folder->GetFile($pdfFileName);
                $outputPath = AsposeApp::$outPutLocation . $pdfFileName;
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
     * Gets the FormField count of the specified PDF document
     */
    public function getFormFieldCount() {
        //build URI
        $strURI = Product::$baseProductUri . '/pdf/' . $this->fileName . '/fields';

        //sign URI
        $signedURI = Utils::sign($strURI);

        //get response stream
        $responseStream = Utils::ProcessCommand($signedURI, 'GET', '');

        $json = json_decode($responseStream);

        return count($json->Fields->List);
    }

    /*
     * Gets the list of FormFields from the specified PDF document
     */
    public function getFormFields() {
        //build URI
        $strURI = Product::$baseProductUri . '/pdf/' . $this->fileName . '/fields';

        //sign URI
        $signedURI = Utils::sign($strURI);

        //get response stream
        $responseStream = Utils::ProcessCommand($signedURI, 'GET', '');

        $json = json_decode($responseStream);

        return $json->Fields->List;
    }

    /*
     * Gets a particular form field
     * $fieldName
     */
    public function getFormField($fieldName) {
        //build URI
        $strURI = Product::$baseProductUri . '/pdf/' . $this->fileName . '/fields/' . $fieldName;

        //sign URI
        $signedURI = Utils::sign($strURI);

        //get response stream
        $responseStream = Utils::ProcessCommand($signedURI, 'GET', '');

        $json = json_decode($responseStream);

        return $json->Field;
    }

    /*
     * Creates an Empty Pdf document
     * @param string $pdfFileName (name of the PDF file to create)
     */
    public function createEmptyPdf($pdfFileName) {
        try {
            //check whether files are set or not
            if ($pdfFileName == '')
                throw new Exception('PDF file name not specified');

            //build URI to create PDF
            $strURI = Product::$baseProductUri . '/pdf/' . $pdfFileName;

            //sign URI
            $signedURI = Utils::sign($strURI);

            $responseStream = Utils::processCommand($signedURI, 'PUT', '', '');

            $v_output = Utils::validateOutput($responseStream);

            if ($v_output === '') {
                //Save PDF file on server
                $folder = new Folder();
                $outputStream = $folder->GetFile($pdfFileName);
                $outputPath = AsposeApp::$outPutLocation . $pdfFileName;
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
     * Adds new page to opened Pdf document
     */
    public function addNewPage() {
        try {
            //check whether files are set or not
            if ($this->fileName == '')
                throw new Exception('PDF file name not specified');

            //build URI to add page
            $strURI = Product::$baseProductUri . '/pdf/' . $this->fileName . '/pages';

            //sign URI
            $signedURI = Utils::sign($strURI);

            $responseStream = Utils::processCommand($signedURI, 'PUT', '', '');

            $v_output = Utils::validateOutput($responseStream);

            if ($v_output === '') {
                //Save PDF file on server
                $folder = new Folder();
                $outputStream = $folder->GetFile($this->fileName);
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
     * Deletes selected page from Pdf document
     * $pageNumber
     */
    public function deletePage($pageNumber) {
        try {
            //check whether files are set or not
            if ($this->fileName == '')
                throw new Exception('PDF file name not specified');

            //build URI to delete page
            $strURI = Product::$baseProductUri . '/pdf/' . $this->fileName . '/pages/' . $pageNumber;

            //sign URI
            $signedURI = Utils::sign($strURI);

            $responseStream = Utils::processCommand($signedURI, 'DELETE', '', '');

            $v_output = Utils::validateOutput($responseStream);

            if ($v_output === '') {
                //Save PDF file on server
                $folder = new Folder();
                $outputStream = $folder->GetFile($this->fileName);
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
     * Moves selected page in Pdf document to new location
     * $pageNumber
     * $newLocation
     */
    public function movePage($pageNumber, $newLocation) {
        try {
            //check whether files are set or not
            if ($this->fileName == '')
                throw new Exception('PDF file name not specified');

            //build URI to move page
            $strURI = Product::$baseProductUri . '/pdf/' . $this->fileName . '/pages/' . $pageNumber .
                    '/movePage?newIndex=' . $newLocation;

            //sign URI
            $signedURI = Utils::sign($strURI);

            $responseStream = Utils::processCommand($signedURI, 'POST', '', '');

            $v_output = Utils::validateOutput($responseStream);

            if ($v_output === '') {
                //Save PDF file on server
                $folder = new Folder();
                $outputStream = $folder->GetFile($this->fileName);
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
     * Replaces Image in PDF File using Local Image Stream
     * $pageNumber
     * $imageIndex
     * $imageStream
     */
    public function replaceImageUsingStream($pageNumber, $imageIndex, $imageStream) {
        try {
            //check whether files are set or not
            if ($this->fileName == '')
                throw new Exception('PDF file name not specified');

            //build URI to replace image
            $strURI = Product::$baseProductUri . '/pdf/' . $this->fileName . '/pages/' . $pageNumber .
                    '/images/' . $imageIndex;

            //sign URI
            $signedURI = Utils::sign($strURI);

            $responseStream = Utils::processCommand($signedURI, 'POST', '', $imageStream);

            $v_output = Utils::validateOutput($responseStream);

            if ($v_output === '') {
                //Save PDF file on server
                $folder = new Folder();
                $outputStream = $folder->GetFile($this->fileName);
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
     * Replaces Image in PDF File using Local Image Stream
     * $pageNumber
     * $imageIndex
     * $fileName
     */
    public function replaceImageUsingFile($pageNumber, $imageIndex, $fileName) {
        try {
            //check whether files are set or not
            if ($this->fileName == '')
                throw new Exception('PDF file name not specified');

            //build URI to replace image
            $strURI = Product::$baseProductUri . '/pdf/' . $this->fileName . '/pages/' . $pageNumber .
                    '/images/' . $imageIndex . '?imageFile=' . $fileName;

            //sign URI
            $signedURI = Utils::sign($strURI);

            $responseStream = Utils::processCommand($signedURI, 'POST', '', '');

            $v_output = Utils::validateOutput($responseStream);

            if ($v_output === '') {
                //Save PDF file on server
                $folder = new Folder();
                $outputStream = $folder->GetFile($this->fileName);
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
     * Get all the properties of the specified document	
     */
    public function getDocumentProperties() {
        try {

            if ($this->fileName == '')
                throw new Exception('PDF file name not specified');

            //build URI to replace image
            $strURI = Product::$baseProductUri . '/pdf/' . $this->fileName . '/documentProperties';

            //sign URI
            $signedURI = Utils::sign($strURI);
            $responseStream = Utils::processCommand($signedURI, 'GET', '', '');

            $response_arr = json_decode($responseStream);

            return $response_arr->DocumentProperties->List;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    /*
     * Get specified properity of the document	
     * @param string $propertyName
     */
    public function getDocumentProperty($propertyName = '') {
        try {

            if ($this->fileName == '')
                throw new Exception('PDF file name not specified');

            if ($propertyName == '')
                throw new Exception('Property name not specified');

            //build URI to replace image
            $strURI = Product::$baseProductUri . '/pdf/' . $this->fileName . '/documentProperties/' . $propertyName;

            //sign URI
            $signedURI = Utils::sign($strURI);
            $responseStream = Utils::processCommand($signedURI, 'GET', '', '');

            $response_arr = json_decode($responseStream);

            return $response_arr->DocumentProperty;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    /*
     * Set specified properity of the document	
     * @param string $propertyName
     * @param string $propertyValue
     */
    public function setDocumentProperty($propertyName = '', $propertyValue = '') {
        try {

            if ($this->fileName == '')
                throw new Exception('PDF file name not specified');

            if ($propertyName == '')
                throw new Exception('Property name not specified');

            //build URI to replace image
            $strURI = Product::$baseProductUri . '/pdf/' . $this->fileName . '/documentProperties/' . $propertyName;

            $putArray['Value'] = $propertyValue;
            $json = json_encode($putArray);

            //sign URI
            $signedURI = Utils::sign($strURI);
            $responseStream = Utils::processCommand($signedURI, 'PUT', 'json', $json);

            $response_arr = json_decode($responseStream);

            return $response_arr->DocumentProperty;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    /*
     * Remove all properties of the document	
     */
    public function removeAllProperties() {
        try {

            if ($this->fileName == '')
                throw new Exception('PDF file name not specified');

            //build URI to replace image
            $strURI = Product::$baseProductUri . '/pdf/' . $this->fileName . '/documentProperties';

            //sign URI
            $signedURI = Utils::sign($strURI);
            $responseStream = Utils::processCommand($signedURI, 'DELETE', '', '');

            $response_arr = json_decode($responseStream);

            return $response_arr->Code == 200 ? true : false;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    public function splitAllPages() {
        try {
            if ($this->fileName == '') {
                throw new Exception('File name not specified');
            }
            $strURI = Product::$baseProductUri . '/pdf/' . $this->fileName . '/split';
            $signedURI = Utils::sign($strURI);
            $responseStream = Utils::processCommand($signedURI, 'POST', '', '');
            $json = json_decode($responseStream);
            $i = 1;
            foreach ($json->Result->Documents as $splitPage) {
                $splitFileName = basename($splitPage->Href);
                $strURI = Product::$baseProductUri . '/storage/file/' . $splitFileName;
                $signedURI = Utils::sign($strURI);
                $responseStream = Utils::processCommand($signedURI, 'GET', '', '');
                $fileName = $this->fileName . '_' . $i . '.pdf';
                $outputFile = AsposeApp::$outPutLocation . $fileName;
                Utils::saveFile($responseStream, $outputFile);
                echo $outputFile.'<br />';
                $i++;
            }
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    public function splitPages($from, $to) {
        try {
            if ($this->fileName == '') {
                throw new Exception('File name not specified');
            }
            $strURI = Product::$baseProductUri . '/pdf/' . $this->fileName . '/split?from=' . $from . '&to=' . $to;
            $signedURI = Utils::sign($strURI);
            $responseStream = Utils::processCommand($signedURI, 'POST', '', '');
            $json = json_decode($responseStream);
            $i = 1;
            foreach ($json->Result->Documents as $splitPage) {
                $splitFileName = basename($splitPage->Href);
                $strURI = Product::$baseProductUri . '/storage/file/' . $splitFileName;
                $signedURI = Utils::sign($strURI);
                $responseStream = Utils::processCommand($signedURI, 'GET', '', '');
                $fileName = $this->fileName . '_' . $i . '.pdf';
                $outputFile = AsposeApp::$outPutLocation . $fileName;
                Utils::saveFile($responseStream, $outputFile);
                echo $outputFile . '<br />';
                $i++;
            }
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    public function splitPagesToAnyFormat($from, $to, $format) {
        try {
            if ($this->fileName == '') {
                throw new Exception('File name not specified');
            }
            $strURI = Product::$baseProductUri . '/pdf/' . $this->fileName . '/split?from=' . $from . '&to=' . $to . '&format=' . $format;
            $signedURI = Utils::sign($strURI);
            $responseStream = Utils::processCommand($signedURI, 'POST', '', '');
            $json = json_decode($responseStream);
            $i = 1;
            foreach ($json->Result->Documents as $splitPage) {
                $splitFileName = basename($splitPage->Href);
                $strURI = Product::$baseProductUri . '/storage/file/' . $splitFileName;
                $signedURI = Utils::sign($strURI);
                $responseStream = Utils::processCommand($signedURI, 'GET', '', '');
                $fileName = $this->fileName . '_' . $i . '.' . $format;
                $outputFile = AsposeApp::$outPutLocation . $fileName;
                Utils::saveFile($responseStream, $outputFile);
                echo $outputFile . '<br />';
                $i++;
            }
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

}