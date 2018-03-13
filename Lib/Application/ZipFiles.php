<?php

namespace Faparicior\ExcelPivotTable\Application;

class ZipFiles
{
    private $fileName;
    /** @var \PclZip $archiver */
    private $archiver;

    /**
     * ZipFiles constructor.
     *
     * @param $fileName
     */
    public function __construct($fileName)
    {
        $this->fileName = $fileName;
        $this->archiver = new \PclZip($fileName);
    }

    /**
     * @param $destination
     */
    public function decomposeExcel($destination)
    {
        $this->archiver->extract($destination);
    }

    /**
     * @param $folder
     */
    public function composeExcel($folder)
    {
        $this->archiver->create($folder.'/', PCLZIP_OPT_REMOVE_PATH, $folder);
    }

    /**
     * Delete temporary folder
     *
     * @param $tmpDir
     */
    public function cleanTemp($tmpDir)
    {
        if (!$tmpDir) return;

        $it = new \RecursiveDirectoryIterator($tmpDir, \RecursiveDirectoryIterator::SKIP_DOTS);
        $files = new \RecursiveIteratorIterator($it,
            \RecursiveIteratorIterator::CHILD_FIRST);
        foreach($files as $file) {
            if ($file->isDir()){
                rmdir($file->getRealPath());
            } else {
                unlink($file->getRealPath());
            }
        }
        rmdir($tmpDir);
    }
}