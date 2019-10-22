<?php

namespace Spatie\SimpleExcel;

use Box\Spout\Common\Entity\Style\Style;
use Box\Spout\Writer\WriterInterface;

class SimpleExcelWriter
{
    /** @var \Box\Spout\Writer\WriterInterface */
    private $writer;

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
            $row = WriterEntityFactory::createRowFromArray($row, $style);
        }

        $this->writer->addRow($row);
    }

    public function __destruct()
    {
        $this->writer->close();
    }
}
