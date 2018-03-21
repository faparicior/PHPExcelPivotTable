<?php

namespace Faparicior\ExcelPivotTable\Domain\Services\ExcelNode;

final class ValueNode implements ExcelNode
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
        return is_numeric($value);
    }

    /**
     * @param $string
     * @param \DOMElement $node
     * @param \DOMDocument $document
     *
     * @return \DOMNode
     */
    public static function valueToExcelNode($string, \DOMElement $node, \DOMDocument $document)
    {
        $node->removeAttribute('t');
        $node->removeAttribute('s');
        $node->nodeValue = '';
        $newNodeV=$document->createElement('v');
        $newNodeV->nodeValue = htmlspecialchars($string);

        return $node->appendChild($newNodeV);
    }
}
