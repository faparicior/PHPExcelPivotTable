<?php

namespace Faparicior\ExcelPivotTable\Domain\Services\ExcelNode;

final class TextNode implements ExcelNode
{
    protected function __construct()
    {

    }

    /**
     * @param $value
     *
     * @return bool
     */
    public static function checkValue($value)
    {
        return is_string($value) || !is_numeric($value);
    }

    /**
     * @param $value
     * @param \DOMElement $node
     * @param \DOMDocument $document
     *
     * @return \DOMNode
     */
    public static function valueToExcelNode($value, \DOMElement $node, \DOMDocument $document)
    {
        $node->setAttribute('t', 'inlineStr');
        $node->removeAttribute('s');
        $node->nodeValue = '';
        $newNodeIs=$document->createElement('is');
        $newNodeT=$document->createElement('t');
        $newNodeT->nodeValue = $value;

        $newNodeIs->appendChild($newNodeT);

        return $node->appendChild($newNodeIs);
    }
}
