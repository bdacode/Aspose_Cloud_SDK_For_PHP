<?php
/*
 * This class contains features to work with text
 */
namespace Aspose\Cloud\Cells;

use Aspose\Cloud\Common\AsposeApp;
use Aspose\Cloud\Common\Utils;
use Aspose\Cloud\Common\Product;
use Aspose\Cloud\Storage\Folder;
use Aspose\Cloud\Exception\AsposeCloudException as Exception;

class TextEditor {

	public $fileName = '';

	public function __construct($fileName) {
		$this->fileName = $fileName;
	}

	/*
	 * Finds a speicif text from Excel document or a worksheet
	 */
	public function findText() {
		$parameters = func_get_args();

		//set parameter values
		if (count($parameters) == 1) {
			$text = $parameters[0];
		} else if (count($parameters) == 2) {
			$WorkSheetName = $parameters[0];
			$text = $parameters[1];
		} else {
			throw new Exception('Invalid number of arguments');
		}
		try {
			//check whether file is set or not
			if ($this->fileName == '')
				throw new Exception('No file name specified');
			$strURI = Product::$baseProductUri . '/cells/' . $this->fileName . ((count($parameters) == 2) ? '/worksheets/' . $WorkSheetName : '') . '/findText?text=' . $text;
			$signedURI = Utils::sign($strURI);
			$responseStream = Utils::processCommand($signedURI, 'POST', '', '');
			$json = json_decode($responseStream);
			return $json -> TextItems -> TextItemList;
		} catch (Exception $e) {
			throw new Exception($e -> getMessage());
		}
	}

	/*
	 * Gets text items from the whole Excel file or a specific worksheet
	 */
	public function getTextItems() {
		$parameters = func_get_args();
		//set parameter values
		if (count($parameters) > 0) {
			$worksheetName = $parameters[0];
		}
		try {
			//check whether file is set or not
			if ($this->fileName == '')
				throw new Exception('No file name specified');
			$strURI = Product::$baseProductUri . '/cells/' . $this->fileName . ((isset($parameters[0])) ? '/worksheets/' . $worksheetName . '/textItems' : '/textItems');
			$signedURI = Utils::sign($strURI);
			$responseStream = Utils::processCommand($signedURI, 'GET', '', '');
			$json = json_decode($responseStream);
			return $json -> TextItems -> TextItemList;
		} catch (Exception $e) {
			throw new Exception($e -> getMessage());
		}
	}

	/*
	 * Replaces all instances of old text with new text in the Excel document or a particular worksheet
	 * @param string $oldText
	 * @param string $newText
	 */
	public function replaceText() {
		$parameters = func_get_args();
		//set parameter values
		if (count($parameters) == 2) {
			$oldText = $parameters[0];
			$newText = $parameters[1];
		} else if (count($parameters) == 3) {
			$oldText = $parameters[0];
			$newText = $parameters[1];
			$worksheetName = $parameters[2];
		} else
			throw new Exception('Invalid number of arguments');
		try {
			//check whether file is set or not
			if ($this->fileName == '')
				throw new Exception('No file name specified');
			//Build URI to replace text
			$strURI = Product::$baseProductUri . '/cells/' . $this->fileName . ((count($parameters) == 3) ? '/worksheets/' . $worksheetName : '') . '/replaceText?oldValue=' . $oldText . '&newValue=' . $newText;
			$signedURI = Utils::sign($strURI);
			$responseStream = Utils::processCommand($signedURI, 'POST', '', '');
			$v_output = Utils::validateOutput($responseStream);
			if ($v_output === '') {
				//Save doc on server
				$folder = new Folder();
				$outputStream = $folder -> GetFile($this->fileName);
				$outputPath = AsposeApp::$outPutLocation . $this->fileName;
				Utils::saveFile($outputStream, $outputPath);
				return $outputPath;
			} else
				return $v_output;
		} catch (Exception $e) {
			throw new Exception($e -> getMessage());
		}
	}

}
