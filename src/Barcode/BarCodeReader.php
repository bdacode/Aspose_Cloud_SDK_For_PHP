<?php
/*
 * reads barcodes from images
 */
namespace Aspose\Cloud\Barcode;

use Aspose\Cloud\Common\Utils;
use Aspose\Cloud\Common\Product;
use Aspose\Cloud\Storage\Folder;
use Aspose\Cloud\Exception\AsposeCloudException as Exception;

class BarcodeReader {

	public $fileName = '';

	public function __construct($fileName) {
		$this->fileName = $fileName;
	}

	/*
	 * reads all or specific barcodes from images
	 * @param string $symbology
	 */
	public function read($symbology) {
		try {
			//check whether file is set or not
			if ($this->fileName == '')
				throw new Exception('No file name specified');
			//build URI to read barcode
			$strURI = Product::$baseProductUri . '/barcode/' . $this->fileName . '/recognize?' . (!isset($symbology) || trim($symbology) === '' ? 'type=' : 'type=' . $symbology);
			//sign URI
			$signedURI = Utils::sign($strURI);
			//get response stream
			$responseStream = Utils::processCommand($signedURI, 'GET', '', '');

			$json = json_decode($responseStream);

			//returns a list of extracted barcodes
			return $json -> Barcodes;
		} catch (Exception $e) {
			throw new Exception($e -> getMessage());
		}
	}

	public function readR($remoteImageName, $remoteFolder, $readType) {
		try {
			if ($this->fileName == '')
				throw new Exception('No file name specified');
			$uri = $this->uriBuilder($remoteImageName, $remoteFolder, $readType);
			$signedURI = Utils::sign($uri);

			$responseStream = Utils::processCommand($signedURI, 'GET', '', '');
			$json = json_decode($responseStream);
			return $json -> Barcodes;
		} catch (Exception $e) {
			throw new Exception($e -> getMessage());
		}
	}

	public function readFromLocalImage($localImage, $remoteFolder, $barcodeReadType) {
		try {
			if ($this->fileName == '')
				throw new Exception('No file name specified');
			$folder = new Folder();
			$folder -> UploadFile($localImage, $remoteFolder);
			$data = $this->ReadR(basename($localImage), $remoteFolder, $barcodeReadType);
			return $data;
		} catch (Exception $e) {
			throw new Exception($e -> getMessage());
			return null;
		}
	}

	public function uriBuilder($remoteImage, $remoteFolder, $readType) {
		$uri = Product::$baseProductUri . '/barcode/';
		if ($remoteImage != null)
			$uri .= $remoteImage . '/';
		$uri .= 'recognize?';
		if ($readType == 'AllSupportedTypes')
			$uri .= 'type=';
		else
			$uri .= 'type=' . $readType;
		if ($remoteFolder != null && trim($remoteFolder) === '')
			$uri .= '&format=' . $remoteFolder;
		if ($remoteFolder != null && trim($remoteFolder) === '')
			$uri .= '&folder=' . $remoteFolder;
		return $uri;
	}

}
