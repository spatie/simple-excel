<?php

namespace Spatie\SimpleExcel;

use Box\Spout\Common\Entity\Style\Style;
use Box\Spout\Writer\Common\Creator\WriterEntityFactory;
use Box\Spout\Writer\WriterInterface;

class SimpleExcelWriter
{
    private WriterInterface $writer;

    private string $path = '';

    private bool $processHeader = true;

    private bool $processingFirstRow = true;

    private int $numberOfRows = 0;

    public static function create(string $file)
    {
        $simpleExcelWriter = new static($file);

        $simpleExcelWriter->getWriter()->openToFile($file);

        return $simpleExcelWriter;
    }

    public static function streamDownload(string $downloadName)
    {
        $simpleExcelWriter = new static($downloadName);

        $simpleExcelWriter->getWriter()->openToBrowser($downloadName);

        return $simpleExcelWriter;
    }

    protected function __construct(string $path)
    {
        $this->writer = WriterEntityFactory::createWriterFromFile($path);

        $this->path = $path;
    }

    public function getPath(): string
    {
        return $this->path;
    }

    public function getWriter(): WriterInterface
    {
        return $this->writer;
    }

    public function getNumberOfRows(): int
    {
        return $this->numberOfRows;
    }

    public function noHeaderRow()
    {
        $this->processHeader = false;

        return $this;
    }

    /**
     * @param \Box\Spout\Common\Entity\Row|array $row
     * @param \Box\Spout\Common\Entity\Style\Style|null $style
     */
    public function addRow($row, Style $style = null)
    {
        if (is_array($row)) {
            if ($this->processHeader && $this->processingFirstRow) {
                $this->writeHeaderFromRow($row);
            }

            $row = WriterEntityFactory::createRowFromArray($row, $style);
        }

        $this->writer->addRow($row);
        $this->numberOfRows++;

        $this->processingFirstRow = false;

        return $this;
    }

    public function addRows(iterable $rows)
    {
        foreach ($rows as $row) {
            $this->addRow($row);
        }
        
        return $this;
    }

    protected function writeHeaderFromRow(array $row)
    {
        $headerValues = array_keys($row);

        $headerRow = WriterEntityFactory::createRowFromArray($headerValues);

        $this->writer->addRow($headerRow);
        $this->numberOfRows++;
    }

    public function toBrowser()
    {
        $this->writer->close();

        exit;
    }

    public function close()
    {
        $this->writer->close();
    }

    public function useDelimiter(string $delimiter): self
    {
        $this->writer->setFieldDelimiter($delimiter);

        return $this;
    }

    public function __destruct()
    {
        $this->close();
    }
}
