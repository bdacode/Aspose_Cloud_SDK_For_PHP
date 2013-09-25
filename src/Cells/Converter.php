<?php
/*
 * converts pages or document into different formats
 */
namespace Aspose\Cloud\Cells;

use Aspose\Cloud\Common\AsposeApp;
use Aspose\Cloud\Common\Utils;
use Aspose\Cloud\Common\Product;
use Aspose\Cloud\Exception\AsposeCloudException as Exception;

class Converter {

	public $fileName = '';
	public $worksheetName = '';
	public $saveFormat = '';

	public function __construct() {
		$parameters = func_get_args();

		//set default values
		if (isset($parameters[0])) {
			$this->fileName = $parameters[0];
		}
		if (isset($parameters[1])) {
			$this->worksheetName = $parameters[1];
		}
		$this->saveFormat = 'xls';
	}

	/*
	 * converts a document to saveformat
	 */
	public function convert() {
		try {
			//check whether file is set or not
			if ($this->fileName == '')
				throw new Exception('No file name specified');
			//Build URI
			$strURI = Product::$baseProductUri . '/cells/' . $this->fileName . '?format=' . $this->saveFormat;
			//Sign URI
			$signedURI = Utils::sign($strURI);
			//Send request and receive response stream
			$responseStream = Utils::processCommand($signedURI, 'GET', '', '');
			//Validate output
			$v_output = Utils::validateOutput($responseStream);
			if ($v_output === '') {
				if ($this->saveFormat == 'html') {
					$saveFormat = 'zip';
				} else {
					$saveFormat = $this->saveFormat;
				}
				$outputPath = Utils::saveFile($responseStream, AsposeApp::$outPutLocation . Utils::getFileName($this->fileName) . '.' . $saveFormat);
				return $outputPath;
			} else {
				return $v_output;
			}
		} catch (Exception $e) {
			throw new Exception($e -> getMessage());
		}
	}

	/*
	 * converts a sheet to image
	 * @param string $worksheetName
	 * @param string $imageFormat
	 */
	public function convertToImage($imageFormat, $worksheetName) {
		try {
			//check whether file and sheet is set or not
			if ($this->fileName == '')
				throw new Exception('No file name specified');
			//Build URI
			$strURI = Product::$baseProductUri . '/cells/' . $this->fileName . '/worksheets/' . $worksheetName . '?format=' . $imageFormat;
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
	 * converts a document to outputFormat
	 * @param string $outputFormat
	 */
	public function save($outputFormat) {
		try {
			//check whether file is set or not
			if ($this->fileName == '')
				throw new Exception('No file name specified');
			//Build URI
			$strURI = Product::$baseProductUri . '/cells/' . $this->fileName . '?format=' . $outputFormat;
			//Sign URI
			$signedURI = Utils::sign($strURI);
			//Send request and receive response stream
			$responseStream = Utils::processCommand($signedURI, 'GET', '', '');
			//Validate output
			$v_output = Utils::validateOutput($responseStream);
			if ($v_output === '') {
				//Save ouput file
				$outputPath = AsposeApp::$outPutLocation . Utils::getFileName($this->fileName) . '.' . $outputFormat;
				Utils::saveFile($responseStream, $outputPath);
				return $outputPath;
			} else
				return $v_output;
		} catch (Exception $e) {
			throw new Exception($e -> getMessage());
		}
	}

	/*
	 * converts a sheet to image
	 * @param string $imageFormat
	 */
	public function worksheetToImage($imageFormat) {
		try {
			//check whether file and sheet is set or not
			if ($this->fileName == '')
				throw new Exception('No file name specified');
			if ($this->worksheetName == '')
				throw new Exception('No worksheet specified');
			//Build URI
			$strURI = Product::$baseProductUri . '/cells/' . $this->fileName . '/worksheets/' . $this->worksheetName . '?format=' . $imageFormat;
			//Sign URI
			$signedURI = Utils::sign($strURI);
			//Send request and receive response stream
			$responseStream = Utils::processCommand($signedURI, 'GET', '', '');
			//Validate output
			$v_output = Utils::validateOutput($responseStream);
			if ($v_output === '') {
				//Save ouput file
				$outputPath = AsposeApp::$outPutLocation . Utils::getFileName($this->fileName) . '_' . $this->worksheetName . '.' . $imageFormat;
				Utils::saveFile($responseStream, $outputPath);
				return $outputPath;
			} else
				return $v_output;
		} catch (Exception $e) {
			throw new Exception($e -> getMessage());
		}
	}

	/*
	 * saves a specific picture from a specific sheet as image
	 * @param $pictureIndex
	 * @param $imageFormat
	 */
	public function pictureToImage($pictureIndex, $imageFormat) {
		try {
			//check whether file and sheet is set or not
			if ($this->fileName == '')
				throw new Exception('No file name specified');
			if ($this->worksheetName == '')
				throw new Exception('No worksheet specified');
			//Build URI
			$strURI = Product::$baseProductUri . '/cells/' . $this->fileName . '/worksheets/' . $this->worksheetName . '/pictures/' . $pictureIndex . '?format=' . $imageFormat;
			//Sign URI
			$signedURI = Utils::sign($strURI);
			//Send request and receive response stream
			$responseStream = Utils::processCommand($signedURI, 'GET', '', '');
			//Validate output
			$v_output = Utils::validateOutput($responseStream);

			if ($v_output === '') {
				//Save ouput file
				$outputPath = AsposeApp::$outPutLocation . Utils::getFileName($this->fileName) . '_' . $this->worksheetName . '.' . $imageFormat;
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
	 * @param $objectIndex
	 * @param $imageFormat
	 */
	public function oleObjectToImage($objectIndex, $imageFormat) {
		try {
			//check whether file and sheet is set or not
			if ($this->fileName == '')
				throw new Exception('No file name specified');
			if ($this->worksheetName == '')
				throw new Exception('No worksheet specified');
			//Build URI
			$strURI = Product::$baseProductUri . '/cells/' . $this->fileName . '/worksheets/' . $this->worksheetName . '/oleobjects/' . $objectIndex . '?format=' . $imageFormat;
			//Sign URI
			$signedURI = Utils::sign($strURI);
			//Send request and receive response stream
			$responseStream = Utils::processCommand($signedURI, 'GET', '', '');
			//Validate output
			$v_output = Utils::validateOutput($responseStream);
			if ($v_output === '') {
				//Save ouput file
				$outputPath = AsposeApp::$outPutLocation . Utils::getFileName($this->fileName) . '_' . $this->worksheetName . '.' . $imageFormat;
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
	 * @param $chartIndex
	 * @param $imageFormat
	 */
	public function chartToImage($chartIndex, $imageFormat) {
		try {
			//check whether file and sheet is set or not
			if ($this->fileName == '')
				throw new Exception('No file name specified');
			if ($this->worksheetName == '')
				throw new Exception('No worksheet specified');
			//Build URI
			$strURI = Product::$baseProductUri . '/cells/' . $this->fileName . '/worksheets/' . $this->worksheetName . '/charts/' . $chartIndex . '?format=' . $imageFormat;
			//Sign URI
			$signedURI = Utils::sign($strURI);
			//Send request and receive response stream
			$responseStream = Utils::processCommand($signedURI, 'GET', '', '');
			//Validate output
			$v_output = Utils::validateOutput($responseStream);
			if ($v_output === '') {
				//Save ouput file
				$outputPath = AsposeApp::$outPutLocation . Utils::getFileName($this->fileName) . '_' . $this->worksheetName . '.' . $imageFormat;
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
	 * @param $shapeIndex
	 * @param $imageFormat
	 */
	public function autoShapeToImage($shapeIndex, $imageFormat) {
		try {
			//check whether file and sheet is set or not
			if ($this->fileName == '')
				throw new Exception('No file name specified');
			if ($this->worksheetName == '')
				throw new Exception('No worksheet specified');
			//Build URI
			$strURI = Product::$baseProductUri . '/cells/' . $this->fileName . '/worksheets/' . $this->worksheetName . '/autoshapes/' . $shapeIndex . '?format=' . $imageFormat;
			//Sign URI
			$signedURI = Utils::sign($strURI);
			//Send request and receive response stream
			$responseStream = Utils::processCommand($signedURI, 'GET', '', '');
			//Validate output
			$v_output = Utils::validateOutput($responseStream);
			if ($v_output === '') {
				//Save ouput file
				$outputPath = AsposeApp::$outPutLocation . Utils::getFileName($this->fileName) . '_' . $this->worksheetName . '.' . $imageFormat;
				Utils::saveFile($responseStream, $outputPath);
				return $outputPath;
			} else
				return $v_output;
		} catch (Exception $e) {
			throw new Exception($e -> getMessage());
		}
	}

	public function convertLocalFile($inputFile, $outputFile, $saveFormat) {
		try {
			if ($inputFile == '') {
				throw new Exception('Please Specify Input File Name along with path');
			}
			if ($outputFile == '') {
				throw new Exception('Please Specify Output File Name along with Extension');
			}
			if ($saveFormat == '') {
				throw new Exception('Please Specify a Save Format');
			}
			$strURI = Product::$baseProductUri . '/cells/convert?format=' . $saveFormat;
			$signedURI = Utils::sign($strURI);
			if (!file_exists($inputFile)) {
				throw new Exception('Input File Doesnt Exists');
			}
			$responseStream = Utils::uploadFileBinary($signedURI, $inputFile, 'xml');
			$v_output = Utils::validateOutput($responseStream);
			if ($v_output === '') {
				if ($saveFormat == 'html') {
					$outputFormat = 'zip';
				} else {
					$outputFormat = $saveFormat;
				}
				if ($outputFile == '') {
					$outputFileName = Utils::getFileName($inputFile) . '.' . $outputFormat;
				} else {
          $outputFileName = Utils::getFileName($outputFile) . '.' . $outputFormat;
        }
				Utils::saveFile($responseStream, AsposeApp::$outPutLocation . $outputFileName);
				return $outputFileName;
			} else {
				return $v_output;
			}
		} catch (Exception $e) {
			throw new Exception($e -> getMessage());
		}
	}

}
