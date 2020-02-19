<?php

namespace Maatwebsite\Excel\Jobs;

use Illuminate\Contracts\Queue\ShouldQueue;
use Illuminate\Foundation\Bus\Dispatchable;
use Maatwebsite\Excel\Concerns\WithMultipleSheets;
use Maatwebsite\Excel\Files\TemporaryFile;
use Maatwebsite\Excel\Writer;

class QueueExport implements ShouldQueue
{
    use ExtendedQueueable, Dispatchable;

    /**
     * @var object
     */
    public $export;

    /**
     * @var string
     */
    private $writerType;

    /**
     * @var TemporaryFile
     */
    private $temporaryFile;

    /**
     * @var int
     */
    public $tries;

    /**
     * @var int
     */
    public $timeout;

    /**
     * @param object        $export
     * @param TemporaryFile $temporaryFile
     * @param string        $writerType
     */
    public function __construct($export, TemporaryFile $temporaryFile, string $writerType)
    {
        $this->export        = $export;
        $this->writerType    = $writerType;
        $this->temporaryFile = $temporaryFile;
        $this->timeout       = $export->timeout ?? null;
        $this->tries         = $export->tries ?? null;
    }

    /**
     * Get the middleware the job should be dispatched through.
     *
     * @return array
     */
    public function middleware()
    {
        return (method_exists($this->export, 'middleware')) ? $this->export->middleware() : [];
    }

    /**
     * @param Writer $writer
     *
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function handle(Writer $writer)
    {
        $writer->open($this->export);

        $sheetExports = [$this->export];
        if ($this->export instanceof WithMultipleSheets) {
            $sheetExports = $this->export->sheets();
        }

        // Pre-create the worksheets
        foreach ($sheetExports as $sheetIndex => $sheetExport) {
            $sheet = $writer->addNewSheet($sheetIndex);
            $sheet->open($sheetExport);
        }

        // Write to temp file with empty sheets.
        $writer->write($sheetExport, $this->temporaryFile, $this->writerType);
    }
}
