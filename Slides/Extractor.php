<?php
/*
 * Extract various types of information from the document
 */
namespace Aspose\Cloud\Slides;

use Aspose\Cloud\Common\Utils;
use Aspose\Cloud\Common\Product;
use Aspose\Cloud\Exception\AsposeCloudException as Exception;

class Extractor {

    public $fileName = '';

    public function __construct($fileName) {
        $this->fileName = $fileName;
    }

    /*
     * Gets total number of images in a presentation
     */
    public function getImageCount($storageName = '', $folder = '') {
        try {
            //check whether file is set or not
            if ($this->fileName == '')
                throw new Exception('No file name specified');

            $strURI = Product::$baseProductUri . '/slides/' . $this->fileName . '/images';
            if ($folder != '') {
                $strURI .= '?folder=' . $folder;
            }
            if ($storageName != '') {
                $strURI .= '&storage=' . $storageName;
            }
            $signedURI = Utils::sign($strURI);

            $responseStream = Utils::processCommand($signedURI, 'GET', '', '');

            $json = json_decode($responseStream);

            return count($json->Images->List);
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    /*
     * Gets number of images in the specified slide
     * @param $slidenumber
     */
    public function getSlideImageCount($slidenumber, $storageName = '', $folder = '') {
        try {
            //check whether file is set or not
            if ($this->fileName == '')
                throw new Exception('No file name specified');

            $strURI = Product::$baseProductUri . '/slides/' . $this->fileName . '/slides/' . $slidenumber . '/images';
            if ($folder != '') {
                $strURI .= '?folder=' . $folder;
            }
            if ($storageName != '') {
                $strURI .= '&storage=' . $storageName;
            }
            $signedURI = Utils::sign($strURI);

            $responseStream = Utils::processCommand($signedURI, 'GET', '', '');

            $json = json_decode($responseStream);

            return count($json->Images->List);
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    /*
     * Gets all shapes from the specified slide
     * @param $slidenumber
     */
    public function getShapes($slidenumber, $storageName = '', $folder = '') {
        try {
            //check whether file is set or not
            if ($this->fileName == '')
                throw new Exception('No file name specified');

            $strURI = Product::$baseProductUri . '/slides/' . $this->fileName . '/slides/' . $slidenumber . '/shapes';
            if ($folder != '') {
                $strURI .= '?folder=' . $folder;
            }
            if ($storageName != '') {
                $strURI .= '&storage=' . $storageName;
            }
            $signedURI = Utils::sign($strURI);

            $responseStream = Utils::processCommand($signedURI, 'GET', '', '');

            $json = json_decode($responseStream);
            
            $shapes = array();

            foreach ($json->ShapeList->ShapesLinks as $shape) {

                $signedURI = Utils::sign($shape->Uri->Href);

                $responseStream = Utils::processCommand($signedURI, 'GET', '', '');

                $json = json_decode($responseStream);

                $shapes[] = $json;
            }

            return $shapes;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    /*
     * Get color scheme from the specified slide
     * $slideNumber
     */
    public function getColorScheme($slideNumber, $storageName = '') {
        try {
            //check whether file is set or not
            if ($this->fileName == '')
                throw new Exception('No file name specified');

            //Build URI to get color scheme
            $strURI = Product::$baseProductUri . '/slides/' . $this->fileName . '/slides/' . $slideNumber . '/theme/colorScheme';
            if ($storageName != '') {
                $strURI .= '?storage=' . $storageName;
            }
            $signedURI = Utils::sign($strURI);

            $responseStream = Utils::processCommand($signedURI, 'GET', '', '');

            $json = json_decode($responseStream);

            return $json->ColorScheme;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    /*
     * Get font scheme from the specified slide
     * $slideNumber
     */
    public function getFontScheme($slideNumber, $storageName = '') {
        try {
            //check whether file is set or not
            if ($this->fileName == '')
                throw new Exception('No file name specified');

            //Build URI to get font scheme
            $strURI = Product::$baseProductUri . '/slides/' . $this->fileName . '/slides/' . $slideNumber . '/theme/fontScheme';
            if ($storageName != '') {
                $strURI .= '?storage=' . $storageName;
            }
            $signedURI = Utils::sign($strURI);

            $responseStream = Utils::processCommand($signedURI, 'GET', '', '');

            $json = json_decode($responseStream);

            return $json->FontScheme;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    /*
     * Get format scheme from the specified slide
     * $slideNumber
     */
    public function getFormatScheme($slideNumber, $storageName = '') {
        try {
            //check whether file is set or not
            if ($this->fileName == '')
                throw new Exception('No file name specified');

            //Build URI to get format scheme
            $strURI = Product::$baseProductUri . '/slides/' . $this->fileName . '/slides/' . $slideNumber . '/theme/formatScheme';
            if ($storageName != '') {
                $strURI .= '?storage=' . $storageName;
            }
            $signedURI = Utils::sign($strURI);

            $responseStream = Utils::processCommand($signedURI, 'GET', '', '');

            $json = json_decode($responseStream);

            return $json->FormatScheme;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    /*
     * Gets placeholder count from a particular slide
     * $slideNumber
     */
    public function getPlaceholderCount($slideNumber, $storageName = '', $folder = '') {
        try {
            //check whether file is set or not
            if ($this->fileName == '')
                throw new Exception('No file name specified');

            $strURI = Product::$baseProductUri . '/slides/' . $this->fileName . '/slides/' . $slideNumber . '/placeholders';
            if ($folder != '') {
                $strURI .= '?folder=' . $folder;
            }
            if ($storageName != '') {
                $strURI .= '&storage=' . $storageName;
            }
            //Build URI to get placeholders
            $signedURI = Utils::sign($strURI);

            $responseStream = Utils::processCommand($signedURI, 'GET', '', '');

            $json = json_decode($responseStream);

            return count($json->Placeholders->PlaceholderLinks);
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    /*
     * Gets a placeholder from a particular slide
     * $slideNumber
     * $placeholderIndex
     */
    public function getPlaceholder($slideNumber, $placeholderIndex, $storageName = '', $folder = '') {
        try {
            //check whether file is set or not
            if ($this->fileName == '')
                throw new Exception('No file name specified');

            $strURI = Product::$baseProductUri . '/slides/' . $this->fileName . '/slides/' . $slideNumber . '/placeholders/' . $placeholderIndex;
            if ($folder != '') {
                $strURI .= '?folder=' . $folder;
            }
            if ($storageName != '') {
                $strURI .= '&storage=' . $storageName;
            }
            //Build URI to get placeholders
            $signedURI = Utils::sign($strURI);

            $responseStream = Utils::processCommand($signedURI, 'GET', '', '');

            $json = json_decode($responseStream);

            return $json->Placeholder;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

}