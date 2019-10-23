<?php

namespace Spatie\SimpleExcel;

use Box\Spout\Writer\WriterInterface;
use Box\Spout\Common\Entity\Style\Style;
use Box\Spout\Writer\Common\Creator\WriterEntityFactory;

class SimpleExcelWriter
{
    /** @var \Box\Spout\Writer\WriterInterface */
    private $writer;

    private $processHeader = true;

    private $processingFirstRow = true;

    public static function create(string $file)
    {
        return new static($file);
    }

    public function __construct(string $path)
    {
        $this->writer = WriterEntityFactory::createWriterFromFile($path);

        $this->writer->openToFile($path);
    }

    public function getWriter(): WriterInterface
    {
        return $this->writer;
    }

    public function noTitleRow()
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

        $this->processingFirstRow = false;

        return $this;
    }

    protected function writeHeaderFromRow(array $row)
    {
        $headerValues = array_keys($row);

        $headerRow = WriterEntityFactory::createRowFromArray($headerValues);

        $this->writer->addRow($headerRow);
    }

    public function __destruct()
    {
        $this->writer->close();
    }
}
