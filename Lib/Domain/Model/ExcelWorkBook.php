<?php

namespace Faparicior\ExcelPivotTable\Domain\Model;

class ExcelWorkBook
{
    const WORKSHEETS_PATH = 'xl/';

    /** @var \DOMDocument $domSheetData */
    private $domSheetData;

    /** @var \DOMDocument $domSheetRel */
    private $domSheetRel;

    /**
     * ExcelWorkBook constructor.
     *
     * @param \DOMDocument $domSheetData
     * @param \DOMDocument $domSheetRel
     */
    public function __construct(\DOMDocument $domSheetData,  \DOMDocument $domSheetRel)
    {
        $this->domSheetData = $domSheetData;
        $this->domSheetRel = $domSheetRel;
    }

    /**
     * @param string $path
     * @param string $workSheetName
     *
     * @return null|\DOMDocument
     */
    public function getWorkBookDataSheet($path, $workSheetName)
    {
        $workSheetFileName = $this->getWorkBookDataSheetName($workSheetName);

        if (!$workSheetFileName) {
            return null;
        }

        $dom = new \DOMDocument();
        $dom->load($path.$workSheetFileName);

        return $dom;
    }

    public function getWorkBookDataSheetName($workSheetName)
    {
        /** @var \DOMElement $element */
        foreach ($this->domSheetData->getElementsByTagName('sheet') as $element)
        {
            if ($element->getAttribute('name') == $workSheetName)
            {
                return self::WORKSHEETS_PATH.$this->getRelationShip($element->getAttribute('r:id'));
            }
        }

        return null;
    }

    /**
     * @param $rId
     * @return string
     */
    private function getRelationShip($rId)
    {
        /** @var \DOMElement $element */
        foreach ($this->domSheetRel->getElementsByTagName('Relationship') as $element) {
            if ($element->getAttribute('Id') == $rId) {
                return $element->getAttribute('Target');
            }
        }
    }

    /**
     * @param string $rangeName
     * @param string $newRange
     *
     * @return bool
     */
    public function modifyRange($rangeName, $newRange)
    {
        $sheetDataRanges = $this->domSheetData->getElementsByTagName('definedName');
        /** @var \DOMElement $dataRange */
        foreach ($sheetDataRanges as $dataRange)
        {
            if($dataRange->getAttribute('name') == $rangeName)
            {
                $dataRangeValue = $dataRange->nodeValue;
                $lastRange = strrpos($dataRangeValue, '$');
                $dataRangeValue = substr($dataRangeValue, 0, $lastRange+1);
                $dataRange->nodeValue = $dataRangeValue.$newRange;

                $this->domSheetData->save($this->domSheetData->documentURI);

                return true;
            }
        }

        return false;
    }
}
