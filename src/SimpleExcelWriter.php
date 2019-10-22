<?php

namespace Spatie\SimpleExcel;

use Box\Spout\Common\Entity\Style\Style;
use Box\Spout\Writer\Common\Creator\WriterEntityFactory;
use Box\Spout\Writer\Common\Creator\WriterFactory;
use Box\Spout\Writer\WriterInterface;

class SimpleExcelWriter
{
    /** @var \Box\Spout\Writer\WriterInterface */
    private $writer;

    private $processHeader = true;

    private $firstRowAdded = false;

    public static function create(string $file)
    {
        return new static($file);
    }

    public function __construct(string $path)
    {
        $this->writer = WriterEntityFactory::createReaderFromFile($path);

        $this->writer->openToFile($path);
    }

    public function getWriter(): WriterInterface
    {
        return $this->writer;
    }

    public function  noHeader()
    {
        $this->processHeader = false;
    }

    /**
     * @param \Box\Spout\Common\Entity\Row|array $row
     * @param \Box\Spout\Common\Entity\Style\Style|null $style
     *
     * @return \Box\Spout\Writer\WriterInterface
     * @throws \Box\Spout\Common\Exception\IOException
     * @throws \Box\Spout\Writer\Exception\WriterNotOpenedException
     */
    public function addRow($row, Style $style = null): WriterInterface
    {
        if (is_array($row)) {
            if (! $this->firstRowAdded) {
                $this->writeHeaderFromRow($row);
            }

            $row = WriterEntityFactory::createRowFromArray($row, $style);
        }

        $this->writer->addRow($row);

        $this->firstRowAdded = true;
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
