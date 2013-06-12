<?php

/**
 *
 * Copyright 2013 Aspose, Inc.
 *
 */

namespace Aspose\Cloud\Common;

use Aspose\Cloud\Exception\AsposeCloudException as Exception;

if (!function_exists('curl_init')) {
    throw new Exception('Aspose needs the CURL PHP extension.');
}
if (!function_exists('json_decode')) {
    throw new Exception('Aspose needs the JSON PHP extension.');
}

/**
 * Provides access to the Aspose Platform.
 *
 * @author Imran Anwar <imran.anwar@Aspose.com>
 */
class Utils {

    public static $http_codes = array(
        100 => 'Continue',
        101 => 'Switching Protocols',
        102 => 'Processing',
        200 => 'OK',
        201 => 'Created',
        202 => 'Accepted',
        203 => 'Non-Authoritative Information',
        204 => 'No Content',
        205 => 'Reset Content',
        206 => 'Partial Content',
        207 => 'Multi-Status',
        300 => 'Multiple Choices',
        301 => 'Moved Permanently',
        302 => 'Found',
        303 => 'See Other',
        304 => 'Not Modified',
        305 => 'Use Proxy',
        306 => 'Switch Proxy',
        307 => 'Temporary Redirect',
        400 => 'Bad Request',
        401 => 'Unauthorized',
        402 => 'Payment Required',
        403 => 'Forbidden',
        404 => 'Not Found',
        405 => 'Method Not Allowed',
        406 => 'Not Acceptable',
        407 => 'Proxy Authentication Required',
        408 => 'Request Timeout',
        409 => 'Conflict',
        410 => 'Gone',
        411 => 'Length Required',
        412 => 'Precondition Failed',
        413 => 'Request Entity Too Large',
        414 => 'Request-URI Too Long',
        415 => 'Unsupported Media Type',
        416 => 'Requested Range Not Satisfiable',
        417 => 'Expectation Failed',
        418 => 'I\'m a teapot',
        422 => 'Unprocessable Entity',
        423 => 'Locked',
        424 => 'Failed Dependency',
        425 => 'Unordered Collection',
        426 => 'Upgrade Required',
        449 => 'Retry With',
        450 => 'Blocked by Windows Parental Controls',
        500 => 'Internal Server Error',
        501 => 'Not Implemented',
        502 => 'Bad Gateway',
        503 => 'Service Unavailable',
        504 => 'Gateway Timeout',
        505 => 'HTTP Version Not Supported',
        506 => 'Variant Also Negotiates',
        507 => 'Insufficient Storage',
        509 => 'Bandwidth Limit Exceeded',
        510 => 'Not Extended'
    );

    /**
     * Performs Aspose Api Request.
     *
     * @param string $url Target Aspose API URL.
     * @param string $method Method to access the API such as GET, POST, PUT and DELETE
     * @param string $headerType XML or JSON
     * @param string $src Post data.
     *
     *
     */
    public static function processCommand($url, $method = 'GET', $headerType = 'XML', $src = '') {

        $method = strtoupper($method);
        $headerType = strtoupper($headerType);
        $session = curl_init();
        curl_setopt($session, CURLOPT_URL, $url);
        if ($method == 'GET') {
            curl_setopt($session, CURLOPT_HTTPGET, 1);
        } else {
            curl_setopt($session, CURLOPT_POST, 1);
            curl_setopt($session, CURLOPT_POSTFIELDS, $src);
            curl_setopt($session, CURLOPT_CUSTOMREQUEST, $method);
        }
        curl_setopt($session, CURLOPT_HEADER, false);
        if ($headerType == 'XML') {
            curl_setopt($session, CURLOPT_HTTPHEADER, array('Accept: application/xml', 'Content-Type: application/xml'));
        } else {
            curl_setopt($session, CURLOPT_HTTPHEADER, array('Content-Type: application/json'));
        }
        curl_setopt($session, CURLOPT_RETURNTRANSFER, true);
        if (preg_match('/^(https)/i', $url))
            curl_setopt($session, CURLOPT_SSL_VERIFYPEER, false);
        $result = curl_exec($session);
        $header = curl_getinfo($session);
        if ($header['http_code'] != 200) {
            throw new Exception('Error Code: ' . $header['http_code'] . ', ' . Utils::$http_codes[$header['http_code']]);
        } else {
            if (preg_match('/You have processed/i', $result) || preg_match('/Your pricing plan allows only/i', $result)) {
                throw new Exception($result);
            }
        }
        curl_close($session);
        return $result;
    }

    /**
     * Performs Aspose Api Request to Upload a file.
     *
     * @param string $url Target Aspose API URL.
     * @param string $localfile Local file 
     * @param string $headerType XML or JSON
     *
     *
     */
    public static function uploadFileBinary($url, $localfile, $headerType = 'XML') {

        $headerType = strtoupper($headerType);
        $fp = fopen($localfile, 'r');
        $session = curl_init();
        curl_setopt($session, CURLOPT_VERBOSE, 1);
        curl_setopt($session, CURLOPT_USERPWD, 'user:password');
        curl_setopt($session, CURLOPT_URL, $url);
        curl_setopt($session, CURLOPT_PUT, 1);
        curl_setopt($session, CURLOPT_RETURNTRANSFER, 1);
        curl_setopt($session, CURLOPT_HEADER, false);
        if ($headerType == 'XML') {
            curl_setopt($session, CURLOPT_HTTPHEADER, array('Accept: application/xml', 'Content-Type: application/xml'));
        } else {
            curl_setopt($session, CURLOPT_HTTPHEADER, array('Content-Type: application/json'));
        }
        curl_setopt($session, CURLOPT_INFILE, $fp);
        curl_setopt($session, CURLOPT_INFILESIZE, filesize($localfile));
        $result = curl_exec($session);
        curl_close($session);
        fclose($fp);
        return $result;
    }

    /**
     * Encode a string to URL-safe base64
     *
     * @param string $value Valure to endode.
     *
     *
     */
    private static function encodeBase64UrlSafe($value) {
        return str_replace(array('+', '/'), array('-', '_'), base64_encode($value));
    }

    /**
     * Decode a string from URL-safe base64
     *
     * @param string $value Value to Decode.
     *
     *
     */
    private static function decodeBase64UrlSafe($value) {
        return base64_decode(str_replace(array('-', '_'), array('+', '/'), $value));
    }

    public static function sign($UrlToSign) {
        // parse the url
        $url = parse_url($UrlToSign);

        if (isset($url['query']) == '')
            $urlPartToSign = $url['path'] . '?appSID=' . AsposeApp::$appSID;
        else
            $urlPartToSign = $url['path'] . '?' . str_replace(' ', '%20', $url['query']) . '&appSID=' . AsposeApp::$appSID;

        // Decode the private key into its binary format
        $decodedKey = self::decodeBase64UrlSafe(AsposeApp::$appKey);

        // Create a signature using the private key and the URL-encoded
        // string using HMAC SHA1. This signature will be binary.
        $signature = hash_hmac('sha1', $urlPartToSign, $decodedKey, true);

        $encodedSignature = self::encodeBase64UrlSafe($signature);

        // return $UrlToSign . '?appSID=' . $this->appSID . '&signature=' . $encodedSignature;
        if (isset($url['query']) == '')
            return $url['scheme'] . '://' . $url['host'] . str_replace(' ', '%20', $url['path']) . '?appSID=' . AsposeApp::$appSID . '&signature=' . $encodedSignature;
        else
            return $url['scheme'] . '://' . $url['host'] . str_replace(' ', '%20', $url['path']) . '?' . str_replace(' ', '%20', $url['query']) . '&appSID=' . AsposeApp::$appSID . '&signature=' . $encodedSignature;
    }

    /**
     * Will get the value of a field in JSON Response
     *
     * @param string $jsonRespose JSON Response string.
     * @param string $fieldName Field to be found.
     *
     * @return getFieldValue($jsonRespose, $fieldName) - String Value of the given Field.
     */
    public function getFieldValue($jsonResponse, $fieldName) {
        return json_decode($jsonResponse)->{$fieldName};
    }

    /**
     * This method parses XML for a count of a particular field.
     *
     * @param string $jsonRespose JSON Response string.
     * @param string $fieldName Field to be found.
     *
     * @return getFieldCount($jsonRespose, $fieldName) - String Value of the given Field.
     */
    public function getFieldCount($jsonResponse, $fieldName) {
        $arr = json_decode($jsonResponse)->{$fieldName};
        return count($arr, COUNT_RECURSIVE);
    }

    /**
     * Copies the contents of input to output. Doesn't close either stream.
     *
     * @param string $input input stream.
     *
     * @return copyStream($input) - Outputs the converted input stream.
     */
    public function copyStream($input) {
        return stream_get_contents($input);
    }

    /**
     * Saves the files
     *
     * @param string $input input stream.
     * @param string $fileName fileName along with the full path.
     *
     *
     */
    public static function saveFile($input, $fileName) {
        $fh = fopen($fileName, 'w') or die('cant open file');
        fwrite($fh, $input);
        fclose($fh);
    }

    public static function getFileName($file) {
        $info = pathinfo($file);
        $file_name = basename($file, '.' . $info['extension']);
        return $file_name;
    }

    public static function validateOutput($result) {
        $string = (string) $result;
        $validate = array('Unknown file format.', 'Unable to read beyond the end of the stream',
            'Index was out of range', 'Cannot read that as a ZipFile', 'Not a Microsoft PowerPoint 2007 presentation',
            'Index was outside the bounds of the array', 'An attempt was made to move the position before the beginning of the stream',
        );
        $invalid = 0;
        foreach ($validate as $key => $value) {
            $pos = strpos($string, $value);
            if ($pos === 1) {
                $invalid = 1;
            }
        }
        if ($invalid == 1)
            return $string;
        else
            return '';
    }

}