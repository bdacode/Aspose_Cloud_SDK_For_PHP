<?php
/*
 * This class contains features to work with charts
 */
namespace Aspose\Cloud\Cells;

use Aspose\Cloud\Common\Utils;
use Aspose\Cloud\Common\Product;
use Aspose\Cloud\Exception\AsposeCloudException as Exception;

class Worksheet {

    public $fileName = '';
    public $worksheetName = '';

    public function __construct($fileName, $worksheetName) {
        $this->fileName = $fileName;
        $this->worksheetName = $worksheetName;
    }

    /*
     * Gets a list of cells
     * $offset
     * $count
     */
    public function getCellsList($offset, $count) {
        try {
            //check whether file is set or not
            if ($this->fileName == '')
                throw new Exception('No file name specified');
            //check whether workshett name is set or not
            if ($this->worksheetName == '')
                throw new Exception('Worksheet name not specified');
            $strURI = Product::$baseProductUri . '/cells/' . $this->fileName .
                    '/worksheets/' . $this->worksheetName . '/cells?offset=' .
                    $offset . '&count=' . $count;
            $signedURI = Utils::sign($strURI);
            $responseStream = Utils::processCommand($signedURI, 'GET', '', '');
            $json = json_decode($responseStream);
            $listCells = array();
            foreach ($json->Cells->CellList as $cell) {
                $strURI = Product::$baseProductUri . '/cells/' . $this->fileName .
                        '/worksheets/' . $this->worksheetName . '/cells' . $cell->link->Href;
                $signedURI = Utils::sign($strURI);
                $responseStream = Utils::processCommand($signedURI, 'GET', '', '');
                $json = json_decode($responseStream);
                array_push($listCells, $json->Cell);
            }
            return $listCells;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    /*
     * Gets a list of rows from the worksheet
     */
    public function getRowsList() {
        try {
            //check whether file is set or not
            if ($this->fileName == '')
                throw new Exception('No file name specified');
            //check whether workshett name is set or not
            if ($this->worksheetName == '')
                throw new Exception('Worksheet name not specified');
            $strURI = Product::$baseProductUri . '/cells/' . $this->fileName .
                    '/worksheets/' . $this->worksheetName . '/cells/rows';
            $signedURI = Utils::sign($strURI);
            $responseStream = Utils::processCommand($signedURI, 'GET', '', '');
            $json = json_decode($responseStream);
            $listRows = array();
            foreach ($json->Rows->RowsList as $row) {
                $strURI = Product::$baseProductUri . '/cells/' . $this->fileName .
                        '/worksheets/' . $this->worksheetName . '/cells/rows' . $row->link->Href;
                $signedURI = Utils::sign($strURI);
                $responseStream = Utils::processCommand($signedURI, 'GET', '', '');
                $json = json_decode($responseStream);
                array_push($listRows, $json->Row);
            }
            return $listRows;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    /*
     * Gets a list of columns from the worksheet
     */
    public function getColumnsList() {
        try {
            //check whether file is set or not
            if ($this->fileName == '')
                throw new Exception('No file name specified');
            //check whether workshett name is set or not
            if ($this->worksheetName == '')
                throw new Exception('Worksheet name not specified');
            $strURI = Product::$baseProductUri . '/cells/' . $this->fileName .
                    '/worksheets/' . $this->worksheetName . '/cells/columns';
            $signedURI = Utils::sign($strURI);
            $responseStream = Utils::processCommand($signedURI, 'GET', '', '');
            $json = json_decode($responseStream);
            $listColumns = array();
            foreach ($json->Columns->ColumnsList as $column) {
                $strURI = Product::$baseProductUri . '/cells/' . $this->fileName .
                        '/worksheets/' . $this->worksheetName . '/cells/columns' . $column->link->Href;
                $signedURI = Utils::sign($strURI);
                $responseStream = Utils::processCommand($signedURI, 'GET', '', '');
                $json = json_decode($responseStream);
                array_push($listColumns, $json->Column);
            }

            return $listColumns;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    /*
     * Gets maximum column index of cell which contains data or style
     * $offset
     * $count
     */
    public function getMaxColumn($offset, $count) {
        try {
            //check whether file is set or not
            if ($this->fileName == '')
                throw new Exception('No file name specified');
            //check whether workshett name is set or not
            if ($this->worksheetName == '')
                throw new Exception('Worksheet name not specified');
            $strURI = Product::$baseProductUri . '/cells/' . $this->fileName .
                    '/worksheets/' . $this->worksheetName . '/cells?offset=' .
                    $offset . '&count=' . $count;
            $signedURI = Utils::sign($strURI);
            $responseStream = Utils::processCommand($signedURI, 'GET', '', '');
            $json = json_decode($responseStream);
            return $json->Cells->MaxColumn;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    /*
     * Gets maximum row index of cell which contains data or style
     * $offset
     * $count
     */
    public function getMaxRow($offset, $count) {
        try {
            //check whether file is set or not
            if ($this->fileName == '')
                throw new Exception('No file name specified');
            //check whether workshett name is set or not
            if ($this->worksheetName == '')
                throw new Exception('Worksheet name not specified');
            $strURI = Product::$baseProductUri . '/cells/' . $this->fileName .
                    '/worksheets/' . $this->worksheetName . '/cells?offset=' .
                    $offset . '&count=' . $count;
            $signedURI = Utils::sign($strURI);
            $responseStream = Utils::processCommand($signedURI, 'GET', '', '');
            $json = json_decode($responseStream);
            return $json->Cells->MaxRow;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    /*
     * Gets cell count in the worksheet
     * $offset
     * $count
     */
    public function getCellsCount($offset, $count) {
        try {
            //check whether file is set or not
            if ($this->fileName == '')
                throw new Exception('No file name specified');
            //check whether workshett name is set or not
            if ($this->worksheetName == '')
                throw new Exception('Worksheet name not specified');
            $strURI = Product::$baseProductUri . '/cells/' . $this->fileName .
                    '/worksheets/' . $this->worksheetName . '/cells?offset=' .
                    $offset . '&count=' . $count;
            $signedURI = Utils::sign($strURI);
            $responseStream = Utils::processCommand($signedURI, 'GET', '', '');
            $json = json_decode($responseStream);
            return $json->Cells->CellCount;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    /*
     * Gets AutoShape count in the worksheet
     */
    public function getAutoShapesCount() {
        try {
            //check whether file is set or not
            if ($this->fileName == '')
                throw new Exception('No file name specified');
            //check whether workshett name is set or not
            if ($this->worksheetName == '')
                throw new Exception('Worksheet name not specified');
            $strURI = Product::$baseProductUri . '/cells/' . $this->fileName .
                    '/worksheets/' . $this->worksheetName . '/autoshapes';
            $signedURI = Utils::sign($strURI);
            $responseStream = Utils::processCommand($signedURI, 'GET', '', '');
            $json = json_decode($responseStream);
            return count($json->AutoShapes->AutoShapeList);
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    /*
     * Gets a specific AutoShape from the sheet
     * $index
     */
    public function getAutoShapeByIndex($index) {
        try {
            //check whether file is set or not
            if ($this->fileName == '')
                throw new Exception('No file name specified');
            //check whether workshett name is set or not
            if ($this->worksheetName == '')
                throw new Exception('Worksheet name not specified');
            $strURI = Product::$baseProductUri . '/cells/' . $this->fileName .
                    '/worksheets/' . $this->worksheetName . '/autoshapes/' . $index;
            $signedURI = Utils::sign($strURI);
            $responseStream = Utils::processCommand($signedURI, 'GET', '', '');
            $json = json_decode($responseStream);
            return $json->AutoShape;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    /*
     * Gets charts count in the worksheet
     */
    public function getChartsCount() {
        try {
            //check whether file is set or not
            if ($this->fileName == '')
                throw new Exception('No file name specified');
            //check whether workshett name is set or not
            if ($this->worksheetName == '')
                throw new Exception('Worksheet name not specified');
            $strURI = Product::$baseProductUri . '/cells/' . $this->fileName .
                    '/worksheets/' . $this->worksheetName . '/charts';
            $signedURI = Utils::sign($strURI);
            $responseStream = Utils::processCommand($signedURI, 'GET', '', '');
            $json = json_decode($responseStream);
            return Count($json->Charts->ChartList);
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    /*
     * Gets a specific chart from the sheet
     * $index
     */
    public function getChartByIndex($index) {
        try {
            //check whether file is set or not
            if ($this->fileName == '')
                throw new Exception('No file name specified');
            //check whether workshett name is set or not
            if ($this->worksheetName == '')
                throw new Exception('Worksheet name not specified');
            $strURI = Product::$baseProductUri . '/cells/' . $this->fileName .
                    '/worksheets/' . $this->worksheetName . '/charts/' . $index;
            $signedURI = Utils::sign($strURI);
            $responseStream = Utils::processCommand($signedURI, 'GET', '', '');
            $json = json_decode($responseStream);
            return $json->Chart;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    /*
     * Gets hyperlinks count in the worksheet
     */
    public function getHyperlinksCount() {
        try {
            //check whether file is set or not
            if ($this->fileName == '')
                throw new Exception('No file name specified');
            //check whether workshett name is set or not
            if ($this->worksheetName == '')
                throw new Exception('Worksheet name not specified');
            $strURI = Product::$baseProductUri . '/cells/' . $this->fileName .
                    '/worksheets/' . $this->worksheetName . '/hyperlinks';
            $signedURI = Utils::sign($strURI);
            $responseStream = Utils::processCommand($signedURI, 'GET', '', '');
            $json = json_decode($responseStream);
            return Count($json->Hyperlinks->HyperlinkList);
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    /*
     * Gets a specific hyperlink from the sheet
     * $index
     */
    public function getHyperlinkByIndex($index) {
        try {
            //check whether file is set or not
            if ($this->fileName == '')
                throw new Exception('No file name specified');
            //check whether workshett name is set or not
            if ($this->worksheetName == '')
                throw new Exception('Worksheet name not specified');
            $strURI = Product::$baseProductUri . '/cells/' . $this->fileName .
                    '/worksheets/' . $this->worksheetName . '/hyperlinks/' . $index;
            $signedURI = Utils::sign($strURI);
            $responseStream = Utils::processCommand($signedURI, 'GET', '', '');
            $json = json_decode($responseStream);
            return $json->Hyperlink;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    public function getComment($cellName) {
        try {
            //check whether file is set or not
            if ($this->fileName == '')
                throw new Exception('No file name specified');
            //check whether workshett name is set or not
            if ($this->worksheetName == '')
                throw new Exception('Worksheet name not specified');
            $strURI = Product::$baseProductUri . '/cells/' . $this->fileName .
                    '/worksheets/' . $this->worksheetName . '/comments/' . $cellName;
            $signedURI = Utils::sign($strURI);
            $responseStream = Utils::processCommand($signedURI, 'GET', '', '');
            $json = json_decode($responseStream);
            return $json->Comment->HtmlNote;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    public function getOleObjectByIndex($index) {
        try {
            //check whether file is set or not
            if ($this->fileName == '')
                throw new Exception('No file name specified');
            //check whether workshett name is set or not
            if ($this->worksheetName == '')
                throw new Exception('Worksheet name not specified');
            $strURI = Product::$baseProductUri . '/cells/' . $this->fileName .
                    '/worksheets/' . $this->worksheetName . '/oleobjects/' . $index;
            $signedURI = Utils::sign($strURI);
            $responseStream = Utils::processCommand($signedURI, 'GET', '', '');
            $json = json_decode($responseStream);
            return $json->OleObject;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    public function getPictureByIndex($index) {
        try {
            //check whether file is set or not
            if ($this->fileName == '')
                throw new Exception('No file name specified');
            //check whether worksheet name is set or not
            if ($this->worksheetName == '')
                throw new Exception('Worksheet name not specified');
            $strURI = Product::$baseProductUri . '/cells/' . $this->fileName .
                    '/worksheets/' . $this->worksheetName . '/pictures/' . $index;
            $signedURI = Utils::sign($strURI);
            $responseStream = Utils::processCommand($signedURI, 'GET', '', '');
            $json = json_decode($responseStream);
            return $json->Picture;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    public function getValidationByIndex($index) {
        try {
            //check whether file is set or not
            if ($this->fileName == '')
                throw new Exception('No file name specified');
            //check whether worksheet name is set or not
            if ($this->worksheetName == '')
                throw new Exception('Worksheet name not specified');
            $strURI = Product::$baseProductUri . '/cells/' . $this->fileName .
                    '/worksheets/' . $this->worksheetName . '/validations/' . $index;
            $signedURI = Utils::sign($strURI);
            $responseStream = Utils::processCommand($signedURI, 'GET', '', '');
            $json = json_decode($responseStream);
            return $json->Validation;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    public function getMergedCellByIndex($index) {
        try {
            //check whether file is set or not
            if ($this->fileName == '')
                throw new Exception('No file name specified');
            //check whether workshett name is set or not
            if ($this->worksheetName == '')
                throw new Exception('Worksheet name not specified');
            $strURI = Product::$baseProductUri . '/cells/' . $this->fileName .
                    '/worksheets/' . $this->worksheetName . '/mergedCells/' . $index;
            $signedURI = Utils::sign($strURI);
            $responseStream = Utils::processCommand($signedURI, 'GET', '', '');
            $json = json_decode($responseStream);
            return $json->MergedCell;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    public function getMergedCellsCount() {
        try {
            //check whether file is set or not
            if ($this->fileName == '')
                throw new Exception('No file name specified');
            //check whether workshett name is set or not
            if ($this->worksheetName == '')
                throw new Exception('Worksheet name not specified');
            $strURI = Product::$baseProductUri . '/cells/' . $this->fileName .
                    '/worksheets/' . $this->worksheetName . '/mergedCells';
            $signedURI = Utils::sign($strURI);
            $responseStream = Utils::processCommand($signedURI, 'GET', '', '');
            $json = json_decode($responseStream);
            return $json->MergedCells->Count;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    public function getValidationsCount() {
        try {
            //check whether file is set or not
            if ($this->fileName == '')
                throw new Exception('No file name specified');
            //check whether workshett name is set or not
            if ($this->worksheetName == '')
                throw new Exception('Worksheet name not specified');
            $strURI = Product::$baseProductUri . '/cells/' . $this->fileName .
                    '/worksheets/' . $this->worksheetName . '/validations';
            $signedURI = Utils::sign($strURI);
            $responseStream = Utils::processCommand($signedURI, 'GET', '', '');
            $json = json_decode($responseStream);
            return count($json->Validations->ValidationList);
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    public function getPicturesCount() {
        try {
            //check whether file is set or not
            if ($this->fileName == '')
                throw new Exception('No file name specified');
            //check whether workshett name is set or not
            if ($this->worksheetName == '')
                throw new Exception('Worksheet name not specified');
            $strURI = Product::$baseProductUri . '/cells/' . $this->fileName .
                    '/worksheets/' . $this->worksheetName . '/pictures';
            $signedURI = Utils::sign($strURI);
            $responseStream = Utils::processCommand($signedURI, 'GET', '', '');
            $json = json_decode($responseStream);
            return count($json->Pictures->PictureList);
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    public function getOleObjectsCount() {
        try {
            //check whether file is set or not
            if ($this->fileName == '')
                throw new Exception('No file name specified');
            //check whether workshett name is set or not
            if ($this->worksheetName == '')
                throw new Exception('Worksheet name not specified');
            $strURI = Product::$baseProductUri . '/cells/' . $this->fileName .
                    '/worksheets/' . $this->worksheetName . '/oleobjects';
            $signedURI = Utils::sign($strURI);
            $responseStream = Utils::processCommand($signedURI, 'GET', '', '');
            $json = json_decode($responseStream);
            return count($json->OleObjects->OleObjectList);
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    public function getCommentsCount() {
        try {
            //check whether file is set or not
            if ($this->fileName == '')
                throw new Exception('No file name specified');
            //check whether workshett name is set or not
            if ($this->worksheetName == '')
                throw new Exception('Worksheet name not specified');
            $strURI = Product::$baseProductUri . '/cells/' . $this->fileName .
                    '/worksheets/' . $this->worksheetName . '/comments';
            $signedURI = Utils::sign($strURI);
            $responseStream = Utils::processCommand($signedURI, 'GET', '', '');
            $json = json_decode($responseStream);
            return count($json->Comments->CommentList);
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    public function hideWorksheet() {
        try {
            //check whether file is set or not
            if ($this->fileName == '')
                throw new Exception('No file name specified');
            //check whether worksheet name is set or not
            if ($this->worksheetName == '')
                throw new Exception('Worksheet name not specified');
            $strURI = Product::$baseProductUri . '/cells/' . $this->fileName .
                    '/worksheets/' . $this->worksheetName . '/visible?isVisible=false';
            $signedURI = Utils::sign($strURI);
            $responseStream = Utils::processCommand($signedURI, 'PUT', '', '');
            $json = json_decode($responseStream);
            if ($json->Code == 200)
                return true;
            else
                return false;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    public function unhideWorksheet() {
        try {
            //check whether file is set or not
            if ($this->fileName == '')
                throw new Exception('No file name specified');
            //check whether worksheet name is set or not
            if ($this->worksheetName == '')
                throw new Exception('Worksheet name not specified');
            $strURI = Product::$baseProductUri . '/cells/' . $this->fileName .
                    '/worksheets/' . $this->worksheetName . '/visible?isVisible=true';
            $signedURI = Utils::sign($strURI);
            $responseStream = Utils::processCommand($signedURI, 'PUT', '', '');
            $json = json_decode($responseStream);
            if ($json->Code == 200)
                return true;
            else
                return false;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    public function moveWorksheet($worksheetName, $position) {
        try {
            //check whether file is set or not
            if ($this->fileName == '')
                throw new Exception('No file name specified');
            //check whether worksheet name is set or not
            if ($this->worksheetName == '')
                throw new Exception('Worksheet name not specified');
            $fieldsArray['DestinationWorsheet'] = $worksheetName;
            $fieldsArray['Position'] = $position;
            $jsonData = json_encode($fieldsArray);
            $strURI = Product::$baseProductUri . '/cells/' . $this->fileName .
                    '/worksheets/' . $this->worksheetName . '/position';
            $signedURI = Utils::sign($strURI);
            $responseStream = Utils::processCommand($signedURI, 'POST', 'json', $jsonData);
            $json = json_decode($responseStream);
            if ($json->Code == 200)
                return true;
            else
                return false;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    public function calculateFormula($formula) {
        try {
            //check whether file is set or not
            if ($this->fileName == '')
                throw new Exception('No file name specified');
            //check whether worksheet name is set or not
            if ($this->worksheetName == '')
                throw new Exception('Worksheet name not specified');
            $strURI = Product::$baseProductUri . '/cells/' . $this->fileName .
                    '/worksheets/' . $this->worksheetName . '/formulaResult?formula=' . $formula;
            $signedURI = Utils::sign($strURI);
            $responseStream = Utils::processCommand($signedURI, 'GET', '', '');
            $json = json_decode($responseStream);
            return $json->Value->Value;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    public function setCellValue($cellName, $valueType, $value) {
        try {
            //check whether file is set or not
            if ($this->fileName == '')
                throw new Exception('No file name specified');
            //check whether worksheet name is set or not
            if ($this->worksheetName == '')
                throw new Exception('Worksheet name not specified');
            $strURI = Product::$baseProductUri . '/cells/' . $this->fileName .
                    '/worksheets/' . $this->worksheetName . '/cells/' . $cellName . '?value=' . $value . '&type=' . $valueType;
            $signedURI = Utils::sign($strURI);
            $responseStream = Utils::processCommand($signedURI, 'POST', '', '');
            $json = json_decode($responseStream);
            if ($json->Code == 200)
                return true;
            else
                return false;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    public function getRowsCount($offset, $count) {
        try {
            //check whether file is set or not
            if ($this->fileName == '')
                throw new Exception('No file name specified');
            //check whether worksheet name is set or not
            if ($this->worksheetName == '')
                throw new Exception('Worksheet name not specified');
            $strURI = Product::$baseProductUri . '/cells/' . $this->fileName .
                    '/worksheets/' . $this->worksheetName . '/cells/rows?offset=' . $offset . '&count=' . $count;
            $signedURI = Utils::sign($strURI);
            $responseStream = Utils::processCommand($signedURI, 'GET', '', '');
            $json = json_decode($responseStream);
            return $json->Rows->RowsCount;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    public function getRow($rowIndex) {
        try {
            //check whether file is set or not
            if ($this->fileName == '')
                throw new Exception('No file name specified');
            //check whether worksheet name is set or not
            if ($this->worksheetName == '')
                throw new Exception('Worksheet name not specified');
            $strURI = Product::$baseProductUri . '/cells/' . $this->fileName .
                    '/worksheets/' . $this->worksheetName . '/cells/rows/' . $rowIndex;
            $signedURI = Utils::sign($strURI);
            $responseStream = Utils::processCommand($signedURI, 'GET', '', '');
            $json = json_decode($responseStream);
            return $json->Row;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    public function deleteRow($rowIndex) {
        try {
            //check whether file is set or not
            if ($this->fileName == '')
                throw new Exception('No file name specified');
            //check whether worksheet name is set or not
            if ($this->worksheetName == '')
                throw new Exception('Worksheet name not specified');
            $strURI = Product::$baseProductUri . '/cells/' . $this->fileName .
                    '/worksheets/' . $this->worksheetName . '/cells/rows/' . $rowIndex;
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

    public function getColumn($columnIndex) {
        try {
            //check whether file is set or not
            if ($this->fileName == '')
                throw new Exception('No file name specified');
            //check whether worksheet name is set or not
            if ($this->worksheetName == '')
                throw new Exception('Worksheet name not specified');
            $strURI = Product::$baseProductUri . '/cells/' . $this->fileName .
                    '/worksheets/' . $this->worksheetName . '/cells/columns/' . $columnIndex;
            $signedURI = Utils::sign($strURI);
            $responseStream = Utils::processCommand($signedURI, 'GET', '', '');
            $json = json_decode($responseStream);
            return $json->Column;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    /*
     * param $dataSort is an array
     * indexes of $dataSort
     * boolean $dataSort['CaseSensitive']
     * boolean $dataSort['HasHeaders']
     * int $dataSort['KeyList']['key']
     * string $dataSort['KeyList']['SortOrder']
     * boolean $dataSort['SortLeftToRight']
     */

    public function sortData(array $dataSort, $cellArea = '') {
        try {
            //check whether file is set or not
            if ($this->fileName == '')
                throw new Exception('No file name specified');
            //check whether worksheet name is set or not
            if ($this->worksheetName == '')
                throw new Exception('Worksheet name not specified');
            $strURI = Product::$baseProductUri . '/cells/' . $this->fileName .
                    '/worksheets/' . $this->worksheetName . '/sort?cellArea=' . $cellArea;
            $json_array = json_encode($dataSort);
            $signedURI = Utils::sign($strURI);
            $responseStream = Utils::processCommand($signedURI, 'POST', 'json', $json_array);
            $json = json_decode($responseStream);
            if ($json->Code == 200)
                return true;
            else
                return false;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    public function setCellStyle($cellName, array $style) {
        try {
            //check whether file is set or not
            if ($this->fileName == '')
                throw new Exception('No file name specified');
            //check whether worksheet name is set or not
            if ($this->worksheetName == '')
                throw new Exception('Worksheet name not specified');
            $strURI = Product::$baseProductUri . '/cells/' . $this->fileName .
                    '/worksheets/' . $this->worksheetName . '/cells/' . $cellName . '/style';
            $jsonArray = json_encode($style);
            $signedURI = Utils::sign($strURI);
            $responseStream = Utils::processCommand($signedURI, 'POST', 'json', $jsonArray);
            $json = json_decode($responseStream);
            if ($json->Code == 200)
                return true;
            else
                return false;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    public function getCell($cellName) {
        try {
            if ($this->fileName == '') {
                throw new Exception('No File Name Specified');
            }
            if ($this->worksheetName == '') {
                throw new Exception('No Worksheet Specified');
            }
            $strURI = Product::$baseProductUri . '/cells/' . $this->fileName . '/worksheets/' . $this->worksheetName . '/cells/' . $cellName;
            $signedURI = Utils::sign($strURI);
            $responseStream = Utils::processCommand($signedURI, 'GET', 'json');
            $json = json_decode($responseStream);
            if ($json->Code == 200) {
                return $json->Cell;
            } else {
                return false;
            }
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    public function getCellStyle($cellName) {
        try {
            if ($this->fileName == '') {
                throw new Exception('No File Name Specified');
            }
            if ($this->worksheetName == '') {
                throw new Exception('No Worksheet Specified');
            }
            $strURI = Product::$baseProductUri . '/cells/' . $this->fileName . '/worksheets/' . $this->worksheetName . '/cells/' . $cellName . '/style';
            $signedURI = Utils::sign($strURI);
            $responseStream = Utils::processCommand($signedURI, 'GET', 'json');
            $json = json_decode($responseStream);
            if ($json->Code == 200) {
                return $json->Style;
            } else {
                return false;
            }
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    public function addPicture($picturePath, $pictureLocation, $upperLeftRow = 0, $upperLeftColumn = 0, $lowerRightRow = 0, $lowerRightColumn = 0) {
        try {
            if ($this->fileName == '') {
                throw new Exception('No File Name Specified');
            }
            if ($this->worksheetName == '') {
                throw new Exception('No Worksheet Specified');
            }
            if ($pictureLocation == 'Server' || $pictureLocation == 'server') {
                $strURI = Product::$baseProductUri . '/cells/' . $this->fileName . '/worksheets/' . $this->worksheetName . '/pictures?upperLeftRow=' .
                        $upperLeftRow . '&upperLeftColumn=' . $upperLeftColumn .
                        '&lowerRightRow=' . $lowerRightRow . '&lowerRightColumn=' . $lowerRightColumn .
                        '&picturePath=' . $picturePath;
                $signedURI = Utils::sign($strURI);
                $responseStream = Utils::processCommand($signedURI, 'PUT');
            } else if ($pictureLocation == 'Local' || $pictureLocation == 'local') {
                if (!file_exists($picturePath)) {
                    throw new Exception('File Does not Exists');
                }
                $stream = file_get_contents($picturePath);
                $strURI = Product::$baseProductUri + '/cells/' . $this->fileName . '/worksheets/' . $this->worksheetName . '/pictures?upperLeftRow=' .
                        $upperLeftRow . '&upperLeftColumn=' . $upperLeftColumn .
                        '&lowerRightRow=' . $lowerRightRow . '&lowerRightColumn=' . $lowerRightColumn .
                        '&picturePath=' . $picturePath;
                $signedURI = Utils::sign($strURI);
                $responseStream = Utils::processCommand($signedURI, 'PUT', '', $stream);
            }
            $json = json_decode($responseStream);
            if ($json->Code == 200) {
                return true;
            } else {
                return false;
            }
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

}