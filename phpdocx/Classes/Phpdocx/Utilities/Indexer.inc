<?php

namespace Phpdocx\Utilities;

/**
 * Return information of a DOCX file
 *
 * @category   Phpdocx
 * @package    utilities
 * @copyright  Copyright (c) Narcea Producciones Multimedia S.L.
 *             (http://www.2mdc.com)
 * @license    phpdocx LICENSE
 * @version    2016.12.10
 * @link       http://www.phpdocx.com
 */
class Indexer
{
    /**
     * @var array Stores the file internal structure like body, header and footer
     */
    private $documentStructure;

    /**
     * @var ZipArchive ZIP file
     */
    private $documentZip;

    /**
     * Class constructor
     *
     * @param string $path File path
     */
    public function __construct($path)
    {
        // check if the file exists.
        // Throw an error if not exists
        if (!file_exists($path)) {
            throw new \Exception('The file ' . $path . ' does not exist');
        }

        // import the file content
        $this->documentZip = new \ZipArchive();
        $this->documentZip->open($path);

        // init the document structure array as empty
        $this->documentStructure = array(
            'body' => array(
                'charts' => array(),
                'images' => array(),
                'links' => array(),
                'text' => array(),
            ),
            'comments' => array(
                'images' => array(),
                'links' => array(),
                'text' => array(),
            ),
            'endnotes' => array(
                'images' => array(),
                'links' => array(),
                'text' => array(),
            ),
            'footers' => array(
                'images' => array(),
                'links' => array(),
                'text' => array(),
            ),
            'footnotes' => array(
                'images' => array(),
                'links' => array(),
                'text' => array(),
            ),
            'headers' => array(
                'images' => array(),
                'links' => array(),
                'text' => array(),
            ),
            'properties' => array(
                'core' => array(),
                'custom' => array(),
            ),
        );

        // parse the document
        $this->parse($path);
    }

    /**
     * Return a file as array or JSON
     *
     * @var string $type Output type: 'array' (default), 'json'
     * @return mixed $output
     * @throws Exception If the output type format not supported
     */
    public function getOutput($output = 'array')
    {
        // if the chosen output type is not supported throw an exception
        if (!in_array($output, array('array', 'json'))) {
            throw new \Exception('The output "' . $output . '" is not supported');
        }

        // output the document after index
        return $this->output($output);
    }

    /**
     * Extract chart contents from a XML string
     *
     * @var string $xml XML string
     * @var string $target Content target
     */
    protected function extractCharts($xml, $target)
    {
        // check if the XML is not empty
        if (!empty($xml)) {
            // load XML content
            $contentDOM = new \DOMDocument();
            $contentDOM->loadXML($xml);

            // do a xpath query getting only image tags
            $contentXpath = new \DOMXPath($contentDOM);
            $contentXpath->registerNamespace('rel', 'http://schemas.openxmlformats.org/package/2006/relationships');
            $chartEntries = $contentXpath->query('//rel:Relationship[@Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"]');

            // iterate charts
            foreach ($chartEntries as $chartEntry) {
                $chartContent = $this->documentZip->getFromName('word/' . $chartEntry->getAttribute('Target'));

                // load chart XML content
                $chartDOM = new \DOMDocument();
                $chartDOM->loadXML($chartContent);

                $contentChartXpath = new \DOMXPath($chartDOM);
                $contentChartXpath->registerNamespace('c', 'http://schemas.openxmlformats.org/drawingml/2006/chart');

                // get chart title
                $chartTitleEntries = $contentChartXpath->query('//c:chart/c:title');
                $chartTitle = '';
                foreach ($chartTitleEntries as $chartTitleEntry) {
                    $chartTitle = $chartTitleEntry->nodeValue;
                }

                // get chart sers
                $chartSerEntries = $contentChartXpath->query('//c:chart//c:ser');
                $data = array();
                foreach ($chartSerEntries as $chartSerEntry) {
                    // get chart ser texts
                    $chartTextSerEntries = $contentChartXpath->query('./c:tx//c:v', $chartSerEntry);

                    $chartTextSer = '';
                    if ($chartTextSerEntries->length > 0) {
                        $chartTextSer = $chartTextSerEntries->item(0)->nodeValue;
                    }

                    // get chart cats
                    $chartCats = array();
                    $chartCatsSerEntries = $contentChartXpath->query('./c:cat//c:v', $chartSerEntry);
                    if ($chartCatsSerEntries->length > 0) {
                        foreach ($chartCatsSerEntries as $chartCatsSerEntry) {
                            $chartCats[] = $chartCatsSerEntry->nodeValue;
                        }
                    }

                    // get chart vals
                    $chartVals = array();
                    $chartValsSerEntries = $contentChartXpath->query('./c:val//c:v', $chartSerEntry);
                    if ($chartValsSerEntries->length > 0) {
                        foreach ($chartValsSerEntries as $chartValsSerEntry) {
                            $chartVals[] = $chartValsSerEntry->nodeValue;
                        }
                    }
                    
                    $data[] = array(
                        'text' => $chartTextSer,
                        'cats' => $chartCats,
                        'vals' => $chartVals,
                    );
                }

                $this->documentStructure[$target]['charts'][] = array(
                    'title' => $chartTitle,
                    'data' => $data,
                );
            }
        }
    }

    /**
     * Extract image contents from a XML string
     *
     * @var string $xml XML string
     * @var string $target Content target
     * @var string $content Content XML to get sizes
     */
    protected function extractImages($xml, $target, $contentTarget)
    {
        // check if the XML is not empty
        if (!empty($xml)) {
            // load XML content
            $contentDOM = new \DOMDocument();
            $contentDOM->loadXML($xml);

            // do a xpath query getting only image tags
            $contentXpath = new \DOMXPath($contentDOM);
            $contentXpath->registerNamespace('rel', 'http://schemas.openxmlformats.org/package/2006/relationships');
            $imageEntries = $contentXpath->query('//rel:Relationship[@Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"]');

            // load XML content
            $contentTargetDOM = new \DOMDocument();
            $contentTargetDOM->loadXML($contentTarget);

            // do a xpath query getting only image tags
            $contentTargetXpath = new \DOMXPath($contentTargetDOM);
            $contentTargetXpath->registerNamespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main');
            $contentTargetXpath->registerNamespace('a', 'http://schemas.openxmlformats.org/drawingml/2006/main');
            $contentTargetXpath->registerNamespace('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');
            $contentTargetXpath->registerNamespace('wp', 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing');

            // init size values
            $heightImage = '';
            $widthImage = '';

            // iterate images
            foreach ($imageEntries as $imageEntry) {
                $imageString = $this->documentZip->getFromName('word/' . $imageEntry->getAttribute('Target'));

                // get the size of the image
                $drawingEntries = $contentTargetXpath->query('//w:drawing/*[.//@r:embed="'.$imageEntry->getAttribute('Id').'"]');
                $drawingChildNodes = $drawingEntries->item(0)->childNodes;
                foreach ($drawingChildNodes as $drawingChildNode) {
                    if ($drawingChildNode->tagName == 'wp:extent') {
                        if ($drawingChildNode->hasAttribute('cx')) {
                            $heightImage = $drawingChildNode->getAttribute('cx');
                        }
                        if ($drawingChildNode->hasAttribute('cy')) {
                            $widthImage = $drawingChildNode->getAttribute('cy');
                        }
                    }
                }

                $this->documentStructure[$target]['images'][] = array(
                    'content' => $imageString,
                    'path' => $imageEntry->getAttribute('Target'),
                    'height_word_emus' => $heightImage,
                    'width_word_emus' => $widthImage,
                    'height_word_inches' => $heightImage/914400,
                    'width_word_inches' => $widthImage/914400,
                    'height_word_cms' => $heightImage/360000,
                    'width_word_cms' => $widthImage/360000,
                );
            }
        }
    }

    /**
     * Extract link contents from a XML string
     *
     * @var string $xml XML string
     * @var string $target Content target
     */
    protected function extractLinks($xml, $target)
    {
        // load XML content
        $contentDOM = new \DOMDocument();
        $contentDOM->loadXML($xml);

        // do a global xpath query getting only text tags
        $contentXpath = new \DOMXPath($contentDOM);
        $contentXpath->registerNamespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main');
        $linkEntries = $contentXpath->query('//w:instrText');

        // iterate link contents and extract the URL
        foreach ($linkEntries as $linkEntry) {
            // if empty text avoid adding the content
            if ($linkEntry->textContent == ' ') {
                continue;
            }

            // remove HYPERLINK and " strings
            $content = str_replace('HYPERLINK', '', $linkEntry->textContent);
            $content = str_replace(array('&quot;', '"'), '', $content);
            $this->documentStructure[$target]['links'][] = trim($content);
        }
    }

    /**
     * Extract document properties from a XML string
     *
     * @var string $xml XML string
     * @var string $target Properties target: core, custom
     */
    protected function extractProperties($xml, $target)
    {
        // load XML content
        $contentDOM = new \DOMDocument();
        $contentDOM->loadXML($xml);

        if ($target == 'core') {
            // do a global xpath query getting only text tags
            $contentXpath = new \DOMXPath($contentDOM);
            $contentXpath->registerNamespace('cp', 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties');
            $propertiesEntries = $contentXpath->query('//cp:coreProperties');

            if ($propertiesEntries->item(0)->childNodes->length > 0) {
                foreach ($propertiesEntries->item(0)->childNodes as $propertyEntry) {
                    // if empty text avoid adding the content
                    if ($propertyEntry->textContent == '') {
                        continue;
                    }

                    // get the name of the property
                    $propertyEntryFullName = explode(':', $propertyEntry->tagName);
                    $nameProperty = $propertyEntryFullName[1];

                    $this->documentStructure['properties']['core'][$nameProperty] = trim($propertyEntry->textContent);
                }
            }
        } else if ($target == 'custom') {
            // do a global xpath query getting only property tags
            $contentXpath = new \DOMXPath($contentDOM);
            $contentXpath->registerNamespace('ns', 'http://schemas.openxmlformats.org/officeDocument/2006/custom-properties');
            $propertiesEntries = $contentXpath->query('//ns:Properties//ns:property');

            if ($propertiesEntries->length > 0) {
                foreach ($propertiesEntries as $propertyEntry) {
                    // if empty text avoid adding the content
                    if ($propertyEntry->textContent == '') {
                        continue;
                    }

                    // get the name of the property
                    $nameProperty = $propertyEntry->getAttribute('name');

                    $this->documentStructure['properties']['custom'][$nameProperty] = trim($propertyEntry->textContent);
                }
            }
        }
    }

    /**
     * Extract text contents from a XML string
     *
     * @var string $xml XML string
     * @return string Text content
     */
    protected function extractTexts($xml)
    {
        // load XML content
        $contentDOM = new \DOMDocument();
        $contentDOM->loadXML($xml);

        // do a global xpath query getting only text tags
        $contentXpath = new \DOMXPath($contentDOM);
        $contentXpath->registerNamespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main');
        $textEntries = $contentXpath->query('//w:t');

        // iterate text content and extract text strings. Add a blank space to separate each string
        $content = '';
        foreach ($textEntries as $textEntry) {
            // if empty text avoid adding the content
            if ($textEntry->textContent == ' ') {
                continue;
            }
            $content .= ' ' . $textEntry->textContent;
        }

        return trim($content);
    }

    /**
     * Return a file as array or JSON
     *
     * @var string $type Output type: 'array' (default), 'json'
     * @return mixed $output
     * @throws Exception If the output format if not supported
     */
    protected function output($type = 'array')
    {
        // array as default
        $output = $this->documentStructure;

        // export as the choosen type
        if ($type == 'json') {
            $output = json_encode($output);
        }

        return $output;
    }

    /**
     * Parse a DOCX file
     *
     * @var string $path File path
     */
    private function parse($path)
    {
        // parse the Content_Types
        $contentTypesContent = $this->documentZip->getFromName('[Content_Types].xml');
        $contentTypesXml = simplexml_load_string($contentTypesContent);

        // get the rels extension, rels as default
        $contentTypesDom = dom_import_simplexml($contentTypesXml);
        $contentTypesXpath = new \DOMXPath($contentTypesDom->ownerDocument);
        $contentTypesXpath->registerNamespace('rel', 'http://schemas.openxmlformats.org/package/2006/content-types');
        $relsEntries = $contentTypesXpath->query('//rel:Default[@ContentType="application/vnd.openxmlformats-package.relationships+xml"]');
        $relsExtension = 'rels';
        if (isset($relsEntries[0])) {
            $relsExtension = $relsEntries[0]->getAttribute('Extension');
        }

        // iterate over the Content_Types and add the header, footer and body contents
        foreach ($contentTypesXml->Override as $override) {
            foreach ($override->attributes() as $attribute => $value) {
                // before adding a content remove the first character to get the right file path
                // removing the first slash of each path
                if ($value == 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml') {
                    // body content
                    
                    // get the file content
                    $content = $this->documentZip->getFromName(substr($override->attributes()->PartName, 1));

                    // extract the text content
                    $textContent = $this->extractTexts($content);
                    
                    $this->documentStructure['body']['text'][] .= $textContent;

                    // extract links
                    $this->extractLinks($content, 'body');

                    // extract images from the same file name plus rels extension
                    $relsPath = str_replace('word/', 'word/_rels/', substr($override->attributes()->PartName, 1)) . '.' . $relsExtension;
                    $contentRels = $this->documentZip->getFromName($relsPath);

                    // extract images
                    $this->extractImages($contentRels, 'body', $content);

                    // extract chars
                    $this->extractCharts($contentRels, 'body');
                } else if ($value == 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml') {
                    // headers content
                     
                    // get the file content
                    $content = $this->documentZip->getFromName(substr($override->attributes()->PartName, 1));

                    // extract the text content
                    $textContent = $this->extractTexts($content);

                    $this->documentStructure['headers']['text'][] = $textContent;

                    // extract links
                    $this->extractLinks($content, 'headers');

                    // extract images from the same file name plus rels extension
                    $relsPath = str_replace('word/', 'word/_rels/', substr($override->attributes()->PartName, 1)) . '.' . $relsExtension;
                    $contentRels = $this->documentZip->getFromName($relsPath);

                    // extract images
                    $this->extractImages($contentRels, 'headers', $content);
                } else if ($value == 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml') {
                    // footers content
                    
                    // get the file content
                    $content = $this->documentZip->getFromName(substr($override->attributes()->PartName, 1));

                    // extract the text content
                    $textContent = $this->extractTexts($content);
                    
                    $this->documentStructure['footers']['text'][] .= $textContent;

                    // extract links
                    $this->extractLinks($content, 'footers');

                    // extract images from the same file name plus rels extension
                    $relsPath = str_replace('word/', 'word/_rels/', substr($override->attributes()->PartName, 1)) . '.' . $relsExtension;
                    $contentRels = $this->documentZip->getFromName($relsPath);

                    // extract images
                    $this->extractImages($contentRels, 'footers', $content);
                } else if ($value == 'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml') {
                    // comments content
                    
                    // get the file content
                    $content = $this->documentZip->getFromName(substr($override->attributes()->PartName, 1));

                    // extract the text content
                    $textContent = $this->extractTexts($content);
                    
                    $this->documentStructure['comments']['text'][] .= $textContent;

                    // extract links
                    $this->extractLinks($content, 'comments');

                    // extract images from the same file name plus rels extension
                    $relsPath = str_replace('word/', 'word/_rels/', substr($override->attributes()->PartName, 1)) . '.' . $relsExtension;
                    $contentRels = $this->documentZip->getFromName($relsPath);

                    // extract images
                    $this->extractImages($contentRels, 'comments', $content);
                } else if ($value == 'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml') {
                    // footnotes content
                    
                    // get the file content
                    $content = $this->documentZip->getFromName(substr($override->attributes()->PartName, 1));

                    // extract the text content
                    $textContent = $this->extractTexts($content);
                    
                    $this->documentStructure['footnotes']['text'][] .= $textContent;

                    // extract links
                    $this->extractLinks($content, 'footnotes');

                    // extract images from the same file name plus rels extension
                    $relsPath = str_replace('word/', 'word/_rels/', substr($override->attributes()->PartName, 1)) . '.' . $relsExtension;
                    $contentRels = $this->documentZip->getFromName($relsPath);

                    // extract images
                    $this->extractImages($contentRels, 'footnotes', $content);
                } else if ($value == 'application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml') {
                    // endnotes content
                    
                    // get the file content
                    $content = $this->documentZip->getFromName(substr($override->attributes()->PartName, 1));

                    // extract the text content
                    $textContent = $this->extractTexts($content);
                    
                    $this->documentStructure['endnotes']['text'][] .= $textContent;

                    // extract links
                    $this->extractLinks($content,'endnotes');

                    // extract images from the same file name plus rels extension
                    $relsPath = str_replace('word/', 'word/_rels/', substr($override->attributes()->PartName, 1)) . '.' . $relsExtension;
                    $contentRels = $this->documentZip->getFromName($relsPath);

                    // extract images
                    $this->extractImages($contentRels, 'endnotes',$content);
                } else if ($value == 'application/vnd.openxmlformats-package.core-properties+xml') {
                    // core properties content
                    
                    // get the file content
                    $content = $this->documentZip->getFromName(substr($override->attributes()->PartName, 1));

                    // extract the properties
                    $this->extractProperties($content, 'core');
                } else if ($value == 'application/vnd.openxmlformats-officedocument.custom-properties+xml') {
                    // custom properties content
                    
                    // get the file content
                    $content = $this->documentZip->getFromName(substr($override->attributes()->PartName, 1));

                    // extract the properties
                    $this->extractProperties($content, 'custom');
                }
            }
        }
    }
}