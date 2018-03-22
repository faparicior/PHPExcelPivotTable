<?php

namespace Faparicior\ExcelPivotTable\Domain\Model;

use Faparicior\ExcelPivotTable\Application\ZipFiles;
use Ramsey\Uuid\Uuid;

class ExcelDocument
{
    const WORKBOOK = 'xl/workbook.xml';

    private $excelOriginal;
    private $tmpPath;
    private $tmpId;

    /**
     * ExcelDocument constructor.
     *
     * @param $excelOriginal
     * @param $tmpPath
     */
    public function __construct($excelOriginal, $tmpPath)
    {
        $this->excelOriginal = $excelOriginal;
        $this->tmpPath = $tmpPath;
        $this->tmpId = Uuid::uuid4()->toString();
    }

    /**
     * @param $workSheetName
     * @param $data
     * @param array $range
     */
    public function changeData($workSheetName, $data, $range)
    {
        $workSheetData = $this->getWorkSheet($workSheetName);
        $excelSheetData = new ExcelDataSheet($workSheetData);
        $excelSheetData->clearData();
        $excelSheetData->addLines($data);

        $workBook = $this->getWorkBook();

        foreach ($range as $item) {
            $workBook->modifyRange($item, $excelSheetData->rowCount());
        }
    }

    /**
     * @param $workSheetName
     * @param $row
     * @param $col
     * @param $data
     */
    public function changeCell($workSheetName, $row, $col, $data)
    {
        $workSheetData = $this->getWorkSheet($workSheetName);
        $excelSheetData = new ExcelDataSheet($workSheetData);

        $excelSheetData->changeCell($row, $col, $data);
    }

    /**
     * @param $range
     * @param $data
     */
    public function modifyRange($range, $data)
    {
        $workBook = $this->getWorkBook();
        $workBook->modifyRange($range, $data->rowCount());
    }

    /**
     * @param string $workSheetName
     *
     * @return \DOMDocument|null
     */
    private function getWorkSheet($workSheetName)
    {
        $workbook = $this->getWorkBook();

        return $workbook->getWorkBookDataSheet($this->tmpPath.'/'.$this->tmpId.'/', $workSheetName);
    }

    /**
     * @return ExcelWorkBook
     */
    private function getWorkBook()
    {
        $dom = new \DOMDocument();
        $dom->load( $this->tmpPath.'/'.$this->tmpId.'/'.self::WORKBOOK);

        return new ExcelWorkBook($dom);
    }

    public function openDocument()
    {
        $excel = new ZipFiles($this->excelOriginal);
        $excel->decomposeExcel($this->tmpPath.$this->tmpId);
    }

    public function saveDocument($newExcel)
    {
        $excelNew = new ZipFiles($newExcel);
        $excelNew->composeExcel($this->tmpPath.$this->tmpId);
        $excelNew->cleanTemp($this->tmpPath.$this->tmpId);
    }
}
