<?php
/*
 * converts pages or document into different formats
 */
namespace Aspose\Cloud\Cells;

use Aspose\Cloud\Common\AsposeApp;
use Aspose\Cloud\Common\Utils;
use Aspose\Cloud\Common\Product;
use Aspose\Cloud\Exception\AsposeCloudException as Exception;

class Extractor {

	public $fileName = '';

	public function __construct($fileName) {
		//set default values
		$this->fileName = $fileName;
	}

	/*
	 * saves a specific picture from a specific sheet as image
	 * @param $worksheetName
	 * @param $pictureIndex
	 * @param $imageFormat
	 */
	public function getPicture($worksheetName, $pictureIndex, $imageFormat) {
		try {
			//check whether file and sheet is set or not
			if ($this->fileName == '')
				throw new Exception('No file name specified');
			//Build URI
			$strURI = Product::$baseProductUri . '/cells/' . $this->fileName . '/worksheets/' . $worksheetName . '/pictures/' . $pictureIndex . '?format=' . $imageFormat;
			//Sign URI
			$signedURI = Utils::sign($strURI);
			//Send request and receive response stream
			$responseStream = Utils::processCommand($signedURI, 'GET', '', '');
			//Validate output
			$v_output = Utils::validateOutput($responseStream);
			if ($v_output === '') {
				//Save ouput file
				$outputPath = AsposeApp::$outPutLocation . Utils::getFileName($this->fileName) . '_' . $worksheetName . '.' . $imageFormat;
				Utils::saveFile($responseStream, $outputPath);
				return $outputPath;
			} else
				return $v_output;
		} catch (Exception $e) {
			throw new Exception($e -> getMessage());
		}
	}

	/*
	 * saves a specific OleObject from a specific sheet as image
	 * @param $worksheetName
	 * @param $objectIndex
	 * @param $imageFormat
	 */
	public function getOleObject($worksheetName, $objectIndex, $imageFormat) {
		try {
			//check whether file and sheet is set or not
			if ($this->fileName == '')
				throw new Exception('No file name specified');
			//Build URI
			$strURI = Product::$baseProductUri . '/cells/' . $this->fileName . '/worksheets/' . $worksheetName . '/oleobjects/' . $objectIndex . '?format=' . $imageFormat;
			//Sign URI
			$signedURI = Utils::sign($strURI);
			//Send request and receive response stream
			$responseStream = Utils::processCommand($signedURI, 'GET', '', '');
			//Validate output
			$v_output = Utils::validateOutput($responseStream);
			if ($v_output === '') {
				//Save ouput file
				$outputPath = AsposeApp::$outPutLocation . Utils::getFileName($this->fileName) . '_' . $worksheetName . '.' . $imageFormat;
				Utils::saveFile($responseStream, $outputPath);
				return $outputPath;
			} else
				return $v_output;
		} catch (Exception $e) {
			throw new Exception($e -> getMessage());
		}
	}

	/*
	 * saves a specific chart from a specific sheet as image
	 * @param $worksheetName
	 * @param $chartIndex
	 * @param $imageFormat
	 */
	public function getChart($worksheetName, $chartIndex, $imageFormat) {
		try {
			//check whether file and sheet is set or not
			if ($this->fileName == '')
				throw new Exception('No file name specified');
			//Build URI
			$strURI = Product::$baseProductUri . '/cells/' . $this->fileName . '/worksheets/' . $worksheetName . '/charts/' . $chartIndex . '?format=' . $imageFormat;
			//Sign URI
			$signedURI = Utils::sign($strURI);
			//Send request and receive response stream
			$responseStream = Utils::processCommand($signedURI, 'GET', '', '');
			//Validate output
			$v_output = Utils::validateOutput($responseStream);
			if ($v_output === '') {
				//Save ouput file
				$outputPath = AsposeApp::$outPutLocation . Utils::getFileName($this->fileName) . '_' . $worksheetName . '.' . $imageFormat;
				Utils::saveFile($responseStream, $outputPath);
				return $outputPath;
			} else
				return $v_output;
		} catch (Exception $e) {
			throw new Exception($e -> getMessage());
		}
	}

	/*
	 * saves a specific auto-shape from a specific sheet as image
	 * @param $worksheetName
	 * @param $shapeIndex
	 * @param $imageFormat
	 */
	public function getAutoShape($worksheetName, $shapeIndex, $imageFormat) {
		try {
			//check whether file and sheet is set or not
			if ($this->fileName == '')
				throw new Exception('No file name specified');

			//Build URI
			$strURI = Product::$baseProductUri . '/cells/' . $this->fileName . '/worksheets/' . $worksheetName . '/autoshapes/' . $shapeIndex . '?format=' . $imageFormat;

			//Sign URI
			$signedURI = Utils::sign($strURI);

			//Send request and receive response stream
			$responseStream = Utils::processCommand($signedURI, 'GET', '', '');

			//Validate output
			$v_output = Utils::validateOutput($responseStream);

			if ($v_output === '') {
				//Save ouput file
				$outputPath = AsposeApp::$outPutLocation . Utils::getFileName($this->fileName) . '_' . $worksheetName . '.' . $imageFormat;
				Utils::saveFile($responseStream, $outputPath);
				return $outputPath;
			} else
				return $v_output;
		} catch (Exception $e) {
			throw new Exception($e -> getMessage());
		}
	}

}
