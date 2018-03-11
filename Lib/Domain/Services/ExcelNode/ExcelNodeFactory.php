<?php

namespace Faparicior\ExcelPivotTable\Domain\Services\ExcelNode;

final class ExcelNodeFactory
{
    public function __construct() {}

    /**
     * @param $value
     * @param $col
     * @param $dom
     *
     * @return \DOMNodeList | \DOMNode | null
     */
    public static function getCellNode($value, $col, $dom)
    {
        if (DateNode::checkValue($value)) {
            return DateNode::valueToExcelNode($value, $col, $dom);
        } elseif (TextNode::checkValue($value)) {
            return TextNode::valueToExcelNode($value, $col, $dom);
        } elseif (ValueNode::checkValue($value)) {
            return ValueNode::valueToExcelNode($value, $col, $dom);
        }

        return null;
    }
}
