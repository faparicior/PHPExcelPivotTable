<?php

namespace Faparicior\ExcelPivotTable\Domain\Model;

use Faparicior\ExcelPivotTable\Domain\Services\CloneDomNode;
use Faparicior\ExcelPivotTable\Domain\Services\ExcelNode\ExcelNodeFactory;

class ExcelDataSheet
{
    /** @var \DOMDocument $domSheetData */
    private $domSheetData;

    /** @var \DOMElement $sheetData */
    private $sheetData;

    /** @var \DOMElement $lineDataSheet */
    private $lineDataSheet;

    /** @var \DOMElement $headerDataSheet */
    private $headerDataSheet;

    private $dataRows = 0;

    /**
     * ExcelDataSheet constructor.
     * @param \DOMDocument $domSheetData
     */
    public function __construct(\DOMDocument $domSheetData)
    {
        $this->domSheetData = $domSheetData;
        $this->sheetData = $domSheetData->getElementsByTagName('sheetData')->item(0);

        $this->setHeaderDataSheet();
        $this->setLineDataSheet();

        $this->addHeader();
    }

    private function setHeaderDataSheet()
    {
        $this->headerDataSheet = $this->domSheetData->getElementsByTagName('row')->item(0);
    }

    private function setLineDataSheet()
    {
        $this->lineDataSheet = $this->domSheetData->getElementsByTagName('row')->item(1);
    }

    public function clearData()
    {
        $parentNode = $this->domSheetData->getElementsByTagName('sheetData')->item(0);

       while($parentNode->hasChildNodes())
       {
           $parentNode->removeChild($parentNode->firstChild);
       }
    }

    public function addHeader()
    {
        $this->nextRowCount();
        $this->domSheetData->getElementsByTagName('sheetData')->item(0)->appendChild($this->headerDataSheet);
    }

    /**
     * @param array $data
     */
    public function addLines($data)
    {
        foreach ($data as $line)
        {
            $this->addLine(array_values($line));
        }

        $this->saveData();
    }

    /**
     * @param $row
     * @param $col
     * @param $data
     */
    public function changeCell($row, $col, $data)
    {
        $cells = $this->domSheetData->getElementsByTagName('c');

        for($i=0; $i<$cells->length; $i++)
        {
            if ($cells->item($i)->getAttribute('r') == $col.$row) {
                ExcelNodeFactory::getCellNode($data, $cells->item($i), $this->domSheetData);
                $this->saveData();

                return;
            }
        }
    }

    /**
     * @param array $arrayLine
     */
    public function addLine($arrayLine)
    {
        $rowCount = $this->nextRowCount();

        /** @var \DOMElement $rows */
        $row = CloneDomNode::cloneNode($this->lineDataSheet, $this->domSheetData);
        $cols = $row->getElementsByTagName('c');
        $row->setAttribute('r', $rowCount);

        $index = 0;
        /** @var \DOMElement $col */
        foreach ($cols as $col) {
            $attribute = $col->getAttribute('r');
            $words = preg_replace('/[0-9]+/', '', $attribute);
            $attribute = $words.$rowCount;
            $col->setAttribute('r', $attribute);
            ExcelNodeFactory::getCellNode($arrayLine[$index], $col, $this->domSheetData);
            $index++;
        }

        $this->sheetData->appendChild($row);
    }

    /**
     * @return \DOMDocument
     */
    public function domSheetData()
    {
        return $this->domSheetData;
    }

    private function nextRowCount()
    {
        return ++$this->dataRows;
    }

    public function rowCount()
    {
        return $this->dataRows;
    }

    private function saveData()
    {
        $this->domSheetData->save($this->domSheetData->documentURI);
    }
}
