<?php

namespace Spatie\SimpleExcel;

use OpenSpout\Common\Entity\Style\Style;
use OpenSpout\Writer\Common\Creator\WriterEntityFactory;
use OpenSpout\Writer\WriterInterface;

class SimpleExcelWriter
{
    private WriterInterface $writer;

    private string $path = '';

    private bool $processHeader = true;

    private bool $processingFirstRow = true;

    private int $numberOfRows = 0;

    private $headerStyle = null;

    public static function create(string $file, string $type = '', callable $configureWriter = null)
    {
        $simpleExcelWriter = new static($file, $type);

        $writer = $simpleExcelWriter->getWriter();

        if ($configureWriter) {
            $configureWriter($writer);
        }

        $writer->openToFile($file);

        return $simpleExcelWriter;
    }

    public static function createWithoutBom(string $file, string $type = '')
    {
        return static::create($file, $type, fn ($writer) => $writer->setShouldAddBOM(false));
    }

    public static function streamDownload(string $downloadName, string $type = '', callable $writerCallback = null)
    {
        $simpleExcelWriter = new static($downloadName, $type);

        $writer = $simpleExcelWriter->getWriter();

        if ($writerCallback) {
            $writerCallback($writer);
        }

        $writer->openToBrowser($downloadName);

        return $simpleExcelWriter;
    }

    protected function __construct(string $path, string $type = '')
    {
        $this->path = $path;

        $this->writer = $type ?
            WriterEntityFactory::createWriter($type) :
            WriterEntityFactory::createWriterFromFile($this->path);
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
     *  @param \Box\Spout\Common\Entity\Style\Style $style
     */
    public function setHeaderStyle($style)
    {
        $this->headerStyle = $style;

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

    public function addRows(iterable $rows, Style $style = null)
    {
        foreach ($rows as $row) {
            $this->addRow($row, $style);
        }

        return $this;
    }

    public function addHeader(array $header): self
    {
        $headerRow = WriterEntityFactory::createRowFromArray($header, $this->headerStyle);

        $this->writer->addRow($headerRow);
        $this->numberOfRows++;

        $this->processingFirstRow = false;

        return $this;
    }

    protected function writeHeaderFromRow(array $row)
    {
        $headerValues = array_keys($row);

        $headerRow = WriterEntityFactory::createRowFromArray($headerValues, $this->headerStyle);

        $this->writer->addRow($headerRow);
        $this->numberOfRows++;
    }

    /**
     * Add a new sheet to the workbook.
     *
     * @param  string|null  $name The name of the sheet. If null, the name will be "SheetX" where X is the sheet index.
     *
     * @return $this
     */
    public function addNewSheetAndMakeItCurrent(?string $name = null): self
    {
        $this->writer->addNewSheetAndMakeItCurrent();
        $this->processingFirstRow = true;
        if ($name) {
            $this->nameCurrentSheet($name);
        }

        return $this;
    }

    /**
     * Sets the name for the current sheet.
     *
     * @param  string  $name
     *
     * @return $this
     */
    public function nameCurrentSheet(string $name): self
    {
        $this->writer->getCurrentSheet()->setName($name);

        return $this;
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
