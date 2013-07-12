<?php
/*
 * Deals with Annotations, Bookmarks, Attachments and Links in PDF document
 */
namespace Aspose\Cloud\Pdf;

use Aspose\Cloud\Common\AsposeApp;
use Aspose\Cloud\Common\Utils;
use Aspose\Cloud\Common\Product;
use Aspose\Cloud\Exception\AsposeCloudException as Exception;

class AnnotationEditor {

    public $fileName = '';

    public function __construct($fileName) {
        $this->fileName = $fileName;
    }

    /*
     * Gets number of annotations on a specified document page
     * @param $pageNumber
     */
    public function getAnnotationsCount($pageNumber) {
        try {
            //check whether file is set or not
            if ($this->fileName == '')
                throw new Exception('No file name specified');
            $strURI = Product::$baseProductUri . '/pdf/' . $this->fileName . '/pages/' . $pageNumber . '/annotations';
            $signedURI = Utils::sign($strURI);
            $responseStream = Utils::processCommand($signedURI, 'GET', '', '');
            $json = json_decode($responseStream);
            return count($json->Annotations->List);
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    /*
     * Gets a specfied annotation on a specified document page
     * @param $pageNumber
     * @param $annotationIndex
     */
    public function getAnnotation($pageNumber, $annotationIndex) {
        try {
            //check whether file is set or not
            if ($this->fileName == '')
                throw new Exception('No file name specified');
            $strURI = Product::$baseProductUri . '/pdf/' . $this->fileName . '/pages/' . $pageNumber . '/annotations/' . $annotationIndex;
            $signedURI = Utils::sign($strURI);
            $responseStream = Utils::processCommand($signedURI, 'GET', '', '');
            $json = json_decode($responseStream);
            return $json->Annotation;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    /*
     * Gets list of all the annotations on a specified document page
     * @param $pageNumber
     */
    public function getAllAnnotations($pageNumber) {
        try {
            //check whether file is set or not
            if ($this->fileName == '')
                throw new Exception('No file name specified');
            $iTotalAnnotation = $this->GetAnnotationsCount($pageNumber);
            $listAnnotations = array();
            for ($index = 1; $index <= $iTotalAnnotation; $index++) {
                array_push($listAnnotations, $this->GetAnnotation($pageNumber, $index));
            }
            return $listAnnotations;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    /*
     * Gets total number of Bookmarks in a Pdf document
     */
    public function getBookmarksCount() {
        try {
            //check whether file is set or not
            if ($this->fileName == '')
                throw new Exception('No file name specified');
            $strURI = Product::$baseProductUri . '/pdf/' . $this->fileName . '/bookmarks';
            $signedURI = Utils::sign($strURI);
            $responseStream = Utils::processCommand($signedURI, 'GET', '', '');
            $json = json_decode($responseStream);
            return count($json->Bookmarks->List);
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    /*
     * Gets number of child bookmarks in a specfied parent bookmark
     * @param $parent
     */
    public function getChildBookmarksCount($parent) {
        try {
            //check whether file is set or not
            if ($this->fileName == '')
                throw new Exception('No file name specified');
            $strURI = Product::$baseProductUri . '/pdf/' . $this->fileName . '/bookmarks/' . $parent . '/bookmarks';
            $signedURI = Utils::sign($strURI);
            $responseStream = Utils::processCommand($signedURI, 'GET', '', '');
            $json = json_decode($responseStream);
            return count($json->Bookmarks->List);
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    /*
     * Gets a specfied Bookmark from a PDF document
     * @param $bookmarkIndex
     */
    public function getBookmark($bookmarkIndex) {
        try {
            //check whether file is set or not
            if ($this->fileName == '')
                throw new Exception('No file name specified');
            $strURI = Product::$baseProductUri . '/pdf/' . $this->fileName . '/bookmarks/' . $bookmarkIndex;
            $signedURI = Utils::sign($strURI);
            $responseStream = Utils::processCommand($signedURI, 'GET', '', '');
            $json = json_decode($responseStream);
            return $json->Bookmark;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    /*
     * Gets a specfied child Bookmark for selected parent bookmark in Pdf document
     * @param $parentIndex
     * @param $childIndex
     */
    public function getChildBookmark($parentIndex, $childIndex) {
        try {
            //check whether file is set or not
            if ($this->fileName == '')
                throw new Exception('No file name specified');
            $strURI = Product::$baseProductUri . '/pdf/' . $this->fileName . '/bookmarks/' . $parentIndex . '/bookmarks/' . $childIndex;
            $signedURI = Utils::sign($strURI);
            $responseStream = Utils::processCommand($signedURI, 'GET', '', '');
            $json = json_decode($responseStream);
            return $json->Bookmark;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    /*
     * Checks whether selected bookmark is parent or child Gets a specfied child Bookmark for selected parent bookmark in Pdf document
     * @param $bookmarkIndex
     */
    public function isChildBookmark($bookmarkIndex) {
        try {
            //check whether file is set or not
            if ($this->fileName == '')
                throw new Exception('No file name specified');
            if ($bookmarkIndex === '')
                throw new Exception('bookmark index not specified');
            $strURI = Product::$baseProductUri . '/pdf/' . $this->fileName . '/bookmarks/' . $bookmarkIndex;
            $signedURI = Utils::sign($strURI);
            $responseStream = Utils::processCommand($signedURI, 'GET', '', '');
            $json = json_decode($responseStream);
            return $json->Bookmark;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    /*
     * Gets list of all the Bookmarks in a Pdf document
     */
    public function getAllBookmarks() {
        try {
            //check whether file is set or not
            if ($this->fileName == '')
                throw new Exception('No file name specified');
            $iTotalBookmarks = $this->GetBookmarksCount();
            $listBookmarks = array();
            for ($index = 1; $index <= $iTotalBookmarks; $index++) {
                array_push($listBookmarks, $this->GetBookmark($index));
            }
            return $listBookmarks;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    /*
     * Gets number of attachments in the Pdf document
     */
    public function getAttachmentsCount() {
        try {
            //check whether file is set or not
            if ($this->fileName == '')
                throw new Exception('No file name specified');
            $strURI = Product::$baseProductUri . '/pdf/' . $this->fileName . '/attachments';
            $signedURI = Utils::sign($strURI);
            $responseStream = Utils::processCommand($signedURI, 'GET', '', '');
            $json = json_decode($responseStream);
            return count($json->Attachments->List);
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    /*
     * Gets selected attachment from Pdf document
     * @param $attachmentIndex
     */
    public function getAttachment($attachmentIndex) {
        try {
            //check whether file is set or not
            if ($this->fileName == '')
                throw new Exception('No file name specified');
            $strURI = Product::$baseProductUri . '/pdf/' . $this->fileName . '/attachments/' . $attachmentIndex;
            $signedURI = Utils::sign($strURI);
            $responseStream = Utils::processCommand($signedURI, 'GET', '', '');
            $json = json_decode($responseStream);
            return $json->Attachment;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    /*
     * Gets List of all the attachments in Pdf document
     */
    public function getAllAttachments() {
        try {
            //check whether file is set or not
            if ($this->fileName == '')
                throw new Exception('No file name specified');
            $iTotalAttachments = $this->GetAttachmentsCount();
            $listAttachments = array();
            for ($index = 1; $index <= $iTotalAttachments; $index++) {
                array_push($listAttachments, $this->GetAttachment($index));
            }
            return $listAttachments;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    /*
     * Download the selected attachment from Pdf document
     * @param string $attachmentIndex
     */
    public function downloadAttachment($attachmentIndex) {
        try {
            //check whether files are set or not
            if ($this->fileName == '')
                throw new Exception('PDF file name not specified');
            $fileInformation = $this->GetAttachment($attachmentIndex);
            //build URI to download attachment
            $strURI = Product::$baseProductUri . '/pdf/' . $this->fileName . '/attachments/' . $attachmentIndex . '/download';
            //sign URI
            $signedURI = Utils::sign($strURI);
            $responseStream = Utils::processCommand($signedURI, 'GET', '', '');
            $v_output = Utils::validateOutput($responseStream);
            if ($v_output === '') {
                Utils::saveFile($responseStream, AsposeApp::$outPutLocation . $fileInformation->Name);
                return '';
            }
            else
                return $v_output;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    /*
     * Gets number of links on a specified document page
     * @param $pageNumber
     */
    public function getLinksCount($pageNumber) {
        try {
            //check whether file is set or not
            if ($this->fileName == '')
                throw new Exception('No file name specified');
            $strURI = Product::$baseProductUri . '/pdf/' . $this->fileName . '/pages/' . $pageNumber . '/links';
            $signedURI = Utils::sign($strURI);
            $responseStream = Utils::processCommand($signedURI, 'GET', '', '');
            $json = json_decode($responseStream);
            return count($json->Links->List);
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    /*
     * Gets a specfied link on a specified document page
     * @param $pageNumber
     * @param $linkIndex
     */
    public function getLink($pageNumber, $linkIndex) {
        try {
            //check whether file is set or not
            if ($this->fileName == '')
                throw new Exception('No file name specified');
            $strURI = Product::$baseProductUri . '/pdf/' . $this->fileName . '/pages/' . $pageNumber . '/links/' . $linkIndex;
            $signedURI = Utils::sign($strURI);
            $responseStream = Utils::processCommand($signedURI, 'GET', '', '');
            $json = json_decode($responseStream);
            return $json->Link;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    /*
     * Gets list of all the links on a specified document page
     * @param $pageNumber
     */
    public function getAllLinks($pageNumber) {
        try {
            //check whether file is set or not
            if ($this->fileName == '')
                throw new Exception('No file name specified');
            $iTotalLinks = $this->GetLinksCount($pageNumber);
            $listLinks = array();
            for ($index = 1; $index <= $iTotalLinks; $index++) {
                array_push($listLinks, $this->GetLink($pageNumber, $index));
            }
            return $listLinks;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

}