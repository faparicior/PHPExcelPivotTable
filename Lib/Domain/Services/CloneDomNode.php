<?php

namespace Faparicior\ExcelPivotTable\Domain\Services;

final class CloneDomNode
{
    protected function __construct()
    {

    }

    /**
     * @param \DOMElement $node
     * @param \DOMDocument $doc
     * @return mixed
     */
    public static function cloneNode(\DOMElement $node, \DOMDocument$doc)
    {
        $newNode=$doc->createElement($node->nodeName);

        foreach($node->attributes as $value)
        {
            $newNode->setAttribute($value->nodeName,$value->value);
        }

        if(!$node->childNodes)
            return $newNode;

        foreach($node->childNodes as $child) {
            if($child->nodeName=="#text"){
                $newNode->appendChild($doc->createTextNode($child->nodeValue));
            } else {
                $newNode->appendChild(self::cloneNode($child,$doc));
            }
        }

        return $newNode;
    }
}