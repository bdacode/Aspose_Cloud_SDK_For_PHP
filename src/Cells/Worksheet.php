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

    }

    /*
     * Gets a list of rows from the worksheet
     */
    public function getRowsList() {
        
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

    }

    /*
     * Gets a list of columns from the worksheet
     */
    public function getColumnsList() {
        
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

    }

    /*
     * Gets maximum column index of cell which contains data or style
     * $offset
     * $count
     */
    public function getMaxColumn($offset, $count) {
        
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

    }

    /*
     * Gets maximum row index of cell which contains data or style
     * $offset
     * $count
     */
    public function getMaxRow($offset, $count) {
        
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

    }

    /*
     * Gets cell count in the worksheet
     * $offset
     * $count
     */
    public function getCellsCount($offset, $count) {
        
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

    }

    /*
     * Gets AutoShape count in the worksheet
     */
    public function getAutoShapesCount() {
        
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

    }

    /*
     * Gets a specific AutoShape from the sheet
     * $index
     */
    public function getAutoShapeByIndex($index) {
        
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

    }

    /*
     * Gets charts count in the worksheet
     */
    public function getChartsCount() {
        
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

    }

    /*
     * Gets a specific chart from the sheet
     * $index
     */
    public function getChartByIndex($index) {
        
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

    }

    /*
     * Gets hyperlinks count in the worksheet
     */
    public function getHyperlinksCount() {
        
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

    }

    /*
     * Gets a specific hyperlink from the sheet
     * $index
     */
    public function getHyperlinkByIndex($index) {
        
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

    }

    /*
     * Delete a specific hyperlink from the sheet
     * $index
     */
    public function deleteHyperlinkByIndex($index) {
        
        //check whether file is set or not
        if ($this->fileName == '')
            throw new Exception('No file name specified');
        //check whether workshett name is set or not
        if ($this->worksheetName == '')
            throw new Exception('Worksheet name not specified');
        $strURI = Product::$baseProductUri . '/cells/' . $this->fileName .
            '/worksheets/' . $this->worksheetName . '/hyperlinks/' . $index;
        $signedURI = Utils::sign($strURI);
        $responseStream = Utils::processCommand($signedURI, 'DELETE', '', '');
        $json = json_decode($responseStream);
        if($json->Status == 'OK')
        {
            return true;
        }
        else
        {
            return false;
        }

    }

    public function getComment($cellName) {
        
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

    }

    public function getOleObjectByIndex($index) {
        
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

    }

    public function deleteOleObjectByIndex($index) {
        
        //check whether file is set or not
        if ($this->fileName == '')
            throw new Exception('No file name specified');
        //check whether workshett name is set or not
        if ($this->worksheetName == '')
            throw new Exception('Worksheet name not specified');
        $strURI = Product::$baseProductUri . '/cells/' . $this->fileName .
            '/worksheets/' . $this->worksheetName . '/oleobjects/' . $index;
        $signedURI = Utils::sign($strURI);
        $responseStream = Utils::processCommand($signedURI, 'DELETE', '', '');
        $json = json_decode($responseStream);
        if($json->Code == 200)
            return true;
        else
            return false;

    }

    public function deleteAllOleObject() {
        
        //check whether file is set or not
        if ($this->fileName == '')
            throw new Exception('No file name specified');
        //check whether workshett name is set or not
        if ($this->worksheetName == '')
            throw new Exception('Worksheet name not specified');
        $strURI = Product::$baseProductUri . '/cells/' . $this->fileName .
            '/worksheets/' . $this->worksheetName . '/oleobjects';
        $signedURI = Utils::sign($strURI);
        $responseStream = Utils::processCommand($signedURI, 'DELETE', '', '');
        $json = json_decode($responseStream);
        if($json->Code == 200)
            return true;
        else
            return false;

    }

    public function getPictureByIndex($index) {
        
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

    }

    public function getValidationByIndex($index) {
        
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

    }

    public function getMergedCellByIndex($index) {
        
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

    }

    public function getMergedCellsCount() {
        
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

    }

    public function getValidationsCount() {
        
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

    }

    public function getPicturesCount() {
        
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

    }

    public function getOleObjectsCount() {
        
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

    }

    public function getCommentsCount() {
        
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

    }

    public function hideWorksheet() {
        
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

    }

    public function unhideWorksheet() {
        
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

    }

    public function FreezePanes($row=1,$col=1,$freezedRows=1,$freezedCols=1) {
        
        //check whether file is set or not
        if ($this->fileName == '')
            throw new Exception('No file name specified');
        //check whether worksheet name is set or not
        if ($this->worksheetName == '')
            throw new Exception('Worksheet name not specified');

        $strURI = Product::$baseProductUri . '/cells/' . $this->fileName .
            '/worksheets/' . $this->worksheetName . '/FreezePanes?row=' . $row . '&col=' . $col . '&freezedRows=' . $freezedRows . '&freezedCols=' . $freezedCols;

        $signedURI = Utils::sign($strURI);

        $responseStream = Utils::processCommand($signedURI, 'PUT', '', '');
        $json = json_decode($responseStream);
        if ($json->Code == 200)
            return true;
        else
            return false;

    }

    public function UnfreezePanes() {
        
        //check whether file is set or not
        if ($this->fileName == '')
            throw new Exception('No file name specified');
        //check whether worksheet name is set or not
        if ($this->worksheetName == '')
            throw new Exception('Worksheet name not specified');

        $strURI = Product::$baseProductUri . '/cells/' . $this->fileName .
            '/worksheets/' . $this->worksheetName . '/FreezePanes';

        $signedURI = Utils::sign($strURI);

        $responseStream = Utils::processCommand($signedURI, 'DELETE', '', '');
        $json = json_decode($responseStream);
        if ($json->Code == 200)
            return true;
        else
            return false;

    }

    public function deleteBackgroundImage() {
        
        //check whether file is set or not
        if ($this->fileName == '')
            throw new Exception('No file name specified');
        //check whether worksheet name is set or not
        if ($this->worksheetName == '')
            throw new Exception('Worksheet name not specified');

        $strURI = Product::$baseProductUri . '/cells/' . $this->fileName .
            '/worksheets/' . $this->worksheetName . '/Background';

        $signedURI = Utils::sign($strURI);

        $responseStream = Utils::processCommand($signedURI, 'DELETE', '', '');
        $json = json_decode($responseStream);
        if ($json->Code == 200)
            return true;
        else
            return false;

    }

    public function setBackgroundImage($backgroundImage="") {
        
        //check whether file is set or not
        if ($this->fileName == '')
            throw new Exception('No file name specified');
        //check whether worksheet name is set or not
        if ($this->worksheetName == '')
            throw new Exception('Worksheet name not specified');

        $strURI = Product::$baseProductUri . '/cells/' . $this->fileName .
            '/worksheets/' . $this->worksheetName . '/Background?imageFile=' . $backgroundImage;

        $signedURI = Utils::sign($strURI);

        $responseStream = Utils::processCommand($signedURI, 'PUT', '', '');
        $json = json_decode($responseStream);
        if ($json->Code == 200)
            return true;
        else
            return false;

    }

    public function updateProperties($properties=array()) {
        
        //check whether file is set or not
        if ($this->fileName == '')
            throw new Exception('No file name specified');
        //check whether worksheet name is set or not
        if ($this->worksheetName == '')
            throw new Exception('Worksheet name not specified');

        $strURI = Product::$baseProductUri . '/cells/' . $this->fileName .
            '/workbook/worksheets/' . $this->worksheetName;

        $signedURI = Utils::sign($strURI);

        $json_data = json_encode($properties);

        $responseStream = Utils::processCommand($signedURI, 'POST', 'json', $json_data);
        $json = json_decode($responseStream);
        if ($json->Code == 200)
            return $json->Worksheet;
        else
            return false;

    }

    public function renameWorksheet($newName) {
        
        //check whether file is set or not
        if ($this->fileName == '')
            throw new Exception('No file name specified');
        //check whether worksheet name is set or not
        if ($this->worksheetName == '')
            throw new Exception('Worksheet name not specified');

        $strURI = Product::$baseProductUri . '/cells/' . $this->fileName .
            '/worksheets/' . $this->worksheetName . '/copy?newname=' . $newName;

        $signedURI = Utils::sign($strURI);

        $responseStream = Utils::processCommand($signedURI, 'POST', '', '');
        $json = json_decode($responseStream);
        if ($json->Code == 200)
            return true;
        else
            return false;

    }

    public function copyWorksheet($newWorksheetName) {
        
        //check whether file is set or not
        if ($this->fileName == '')
            throw new Exception('No file name specified');
        //check whether worksheet name is set or not
        if ($this->worksheetName == '')
            throw new Exception('Worksheet name not specified');

        $strURI = Product::$baseProductUri . '/cells/' . $this->fileName .
            '/worksheets/' . $newWorksheetName . '/copy?sourcesheet=' . $this->worksheetName;

        $signedURI = Utils::sign($strURI);

        $responseStream = Utils::processCommand($signedURI, 'POST', '', '');
        $json = json_decode($responseStream);
        if ($json->Code == 200)
            return true;
        else
            return false;

    }

    public function moveWorksheet($worksheetName, $position) {
        
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

    }

    public function calculateFormula($formula) {
        
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

    }

    public function setCellValue($cellName, $valueType, $value) {
        
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

    }

    public function getRowsCount($offset, $count) {
        
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

    }

    public function copyRows($sourceRowIndex=1,$destRowIndex=1,$rowNumber=1) {
        
        //check whether file is set or not
        if ($this->fileName == '')
            throw new Exception('No file name specified');
        //check whether worksheet name is set or not
        if ($this->worksheetName == '')
            throw new Exception('Worksheet name not specified');
        $strURI = Product::$baseProductUri . '/cells/' . $this->fileName .
            '/worksheets/' . $this->worksheetName . '/cells/rows/copy?sourceRowIndex=' . $sourceRowIndex . '&destinationRowIndex=' . $destRowIndex . '&rowNumber=' . $rowNumber;
        $signedURI = Utils::sign($strURI);
        $responseStream = Utils::processCommand($signedURI, 'POST', '', '');
        $json = json_decode($responseStream);
        if ($json->Code == 200)
            return true;
        else
            return false;

    }

    public function autofitRows($firstIndex=1,$lastIndex=1) {
        
        //check whether file is set or not
        if ($this->fileName == '')
            throw new Exception('No file name specified');
        //check whether worksheet name is set or not
        if ($this->worksheetName == '')
            throw new Exception('Worksheet name not specified');
        $strURI = Product::$baseProductUri . '/cells/' . $this->fileName .
            '/worksheets/' . $this->worksheetName . '/cells/rows/autofit?firstIndex=' . $firstIndex . '&lastIndex=' . $lastIndex;
        $signedURI = Utils::sign($strURI);
        $responseStream = Utils::processCommand($signedURI, 'POST', '', '');
        $json = json_decode($responseStream);
        if ($json->Code == 200)
            return true;
        else
            return false;

    }

    public function groupRows($firstIndex=1,$lastIndex=1,$hide=false) {
        
        //check whether file is set or not
        if ($this->fileName == '')
            throw new Exception('No file name specified');
        //check whether worksheet name is set or not
        if ($this->worksheetName == '')
            throw new Exception('Worksheet name not specified');
        $strURI = Product::$baseProductUri . '/cells/' . $this->fileName .
            '/worksheets/' . $this->worksheetName . '/cells/rows/group?firstIndex=' . $firstIndex . '&lastIndex=' . $lastIndex . '&hide=' . $hide;
        $signedURI = Utils::sign($strURI);
        $responseStream = Utils::processCommand($signedURI, 'POST', '', '');
        $json = json_decode($responseStream);
        if ($json->Code == 200)
            return true;
        else
            return false;

    }

    public function ungroupRows($firstIndex=1,$lastIndex=1) {
        
        //check whether file is set or not
        if ($this->fileName == '')
            throw new Exception('No file name specified');
        //check whether worksheet name is set or not
        if ($this->worksheetName == '')
            throw new Exception('Worksheet name not specified');
        $strURI = Product::$baseProductUri . '/cells/' . $this->fileName .
            '/worksheets/' . $this->worksheetName . '/cells/rows/ungroup?firstIndex=' . $firstIndex . '&lastIndex=' . $lastIndex;
        $signedURI = Utils::sign($strURI);
        $responseStream = Utils::processCommand($signedURI, 'POST', '', '');
        $json = json_decode($responseStream);
        if ($json->Code == 200)
            return true;
        else
            return false;

    }

    public function unhideRows($startRow=1,$totalRows=1) {
        
        //check whether file is set or not
        if ($this->fileName == '')
            throw new Exception('No file name specified');
        //check whether worksheet name is set or not
        if ($this->worksheetName == '')
            throw new Exception('Worksheet name not specified');
        $strURI = Product::$baseProductUri . '/cells/' . $this->fileName .
            '/worksheets/' . $this->worksheetName . '/cells/rows/unhide?startrow=' . $startRow . '&totalRows=' . $totalRows;
        $signedURI = Utils::sign($strURI);
        $responseStream = Utils::processCommand($signedURI, 'POST', '', '');
        $json = json_decode($responseStream);
        if ($json->Code == 200)
            return true;
        else
            return false;

    }

    public function hideRows($startRow=1,$totalRows=1) {
        
        //check whether file is set or not
        if ($this->fileName == '')
            throw new Exception('No file name specified');
        //check whether worksheet name is set or not
        if ($this->worksheetName == '')
            throw new Exception('Worksheet name not specified');
        $strURI = Product::$baseProductUri . '/cells/' . $this->fileName .
            '/worksheets/' . $this->worksheetName . '/cells/rows/hide?startrow=' . $startRow . '&totalRows=' . $totalRows;
        $signedURI = Utils::sign($strURI);
        $responseStream = Utils::processCommand($signedURI, 'POST', '', '');
        $json = json_decode($responseStream);
        if ($json->Code == 200)
            return true;
        else
            return false;

    }

    public function getRow($rowIndex) {
        
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

    }

    public function deleteRow($rowIndex) {
        
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

    }

    public function getColumn($columnIndex) {
        
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

    }

    public function setCellStyle($cellName, array $style) {
        
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

    }

    public function getCell($cellName) {
        
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

    }

    public function getCellStyle($cellName) {
        
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

    }

    public function addPicture($picturePath, $pictureLocation, $upperLeftRow = 0, $upperLeftColumn = 0, $lowerRightRow = 0, $lowerRightColumn = 0) {
        
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

    }

    public function addOleObject($oleFile='', $imageFile='', $upperLeftRow = 0, $upperLeftColumn = 0, $height = 0, $width = 0) {
        
        if ($this->fileName == '') {
            throw new Exception('No File Name Specified');
        }
        if ($this->worksheetName == '') {
            throw new Exception('No Worksheet Specified');
        }
        $strURI = Product::$baseProductUri . '/cells/' . $this->fileName . '/worksheets/' . $this->worksheetName . '/oleobjects?upperLeftRow=' .
            $upperLeftRow . '&upperLeftColumn=' . $upperLeftColumn .
            '&height=' . $height . '&width=' . $width .
            '&oleFile=' . $oleFile . '&imageFile='.$imageFile;
        $signedURI = Utils::sign($strURI);

        $responseStream = Utils::processCommand($signedURI, 'PUT');
        $json = json_decode($responseStream);
        if ($json->Code == 200) {
            return true;
        } else {
            return false;
        }

    }

    /*
	 * Update a specific object from on specific sheet
	 * @param $objectIndex
	 */

    public function updateOleObject($objectIndex,$object_data) {
        
        //check whether file and sheet is set or not
        if ($this->fileName == '')
            throw new Exception('No file name specified');
        //Build URI
        $strURI = Product::$baseProductUri . '/cells/' . $this->fileName . '/worksheets/' . $this->worksheetName . '/oleobjects/' . $objectIndex;
        //Sign URI
        $signedURI = Utils::sign($strURI);
        //Send request and receive response stream
        $responseStream = Utils::processCommand($signedURI, 'POST', 'json', $object_data);
        $json = json_decode($responseStream);
        return $json;

    }

    /*
	 * Update a specific picture from on specific sheet
	 * @param $pictureIndex
	 */
    public function updatePicture($pictureIndex,$picture_data) {
        
        //check whether file and sheet is set or not
        if ($this->fileName == '')
            throw new Exception('No file name specified');
        //Build URI
        $strURI = Product::$baseProductUri . '/cells/' . $this->fileName . '/worksheets/' . $this->worksheetName . '/pictures/' . $pictureIndex;
        //Sign URI
        $signedURI = Utils::sign($strURI);
        //Send request and receive response stream
        $responseStream = Utils::processCommand($signedURI, 'POST', 'json', $picture_data);
        $json = json_decode($responseStream);
        return $json;

    }

    /*
	 * Delete a specific picture from a specific sheet
	 * @param $pictureIndex
	 */
    public function deletePicture($pictureIndex) {
        
        //check whether file and sheet is set or not
        if ($this->fileName == '')
            throw new Exception('No file name specified');
        //Build URI
        $strURI = Product::$baseProductUri . '/cells/' . $this->fileName . '/worksheets/' . $this->worksheetName . '/pictures/' . $pictureIndex;
        //Sign URI
        $signedURI = Utils::sign($strURI);
        //Send request and receive response stream
        $responseStream = Utils::processCommand($signedURI, 'DELETE', '', '');
        $json = json_decode($responseStream);
        if ($json->Code == 200) {
            return true;
        } else {
            return false;
        }

    }

    /*
	 * Delete all pictures from a specific sheet
	 */
    public function deleteAllPictures() {
        
        //check whether file and sheet is set or not
        if ($this->fileName == '')
            throw new Exception('No file name specified');
        //Build URI
        $strURI = Product::$baseProductUri . '/cells/' . $this->fileName . '/worksheets/' . $this->worksheetName . '/pictures';
        //Sign URI
        $signedURI = Utils::sign($strURI);
        //Send request and receive response stream
        $responseStream = Utils::processCommand($signedURI, 'DELETE', '', '');
        $json = json_decode($responseStream);
        if ($json->Code == 200) {
            return true;
        } else {
            return false;
        }

    }

}