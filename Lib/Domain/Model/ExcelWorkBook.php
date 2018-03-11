<?php

namespace Faparicior\ExcelPivotTable\Domain\Model;

class ExcelWorkBook
{
    const WORKSHEETS_PATH = 'xl/worksheets/';

    /** @var \DOMDocument $domSheetData */
    private $domSheetData;

    /**
     * ExcelWorkBook constructor.
     *
     * @param \DOMDocument $domSheetData
     */
    public function __construct(\DOMDocument $domSheetData)
    {
        $this->domSheetData = $domSheetData;
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
                $sheetId = $element->getAttribute('sheetId');
                return self::WORKSHEETS_PATH.'sheet'.$sheetId.'.xml';
            }
        }

        return null;
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
