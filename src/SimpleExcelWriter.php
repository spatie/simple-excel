<?php

namespace Spatie\SimpleExcel;

use OpenSpout\Common\Entity\Row;
use OpenSpout\Common\Entity\Style\Style;
use OpenSpout\Writer\CSV\Options as CSVOptions;
use OpenSpout\Writer\CSV\Writer;
use OpenSpout\Writer\WriterInterface;

class SimpleExcelWriter
{
    private WriterInterface $writer;

    private string $path = '';

    private bool $processHeader = true;

    private bool $processingFirstRow = true;

    private int $numberOfRows = 0;

    private ?Style $headerStyle = null;

    protected CSVOptions $csvOptions;

    public static function create(
        string $file,
        string $type = '',
        callable $configureWriter = null,
        ?string $delimiter = null,
        ?bool $shouldAddBom = null,
    ) {
        $simpleExcelWriter = new static(
            path: $file,
            type: $type,
            delimiter: $delimiter,
            shouldAddBom: $shouldAddBom,
        );

        $writer = $simpleExcelWriter->getWriter();

        if ($configureWriter) {
            $configureWriter($writer);
        }

        $writer->openToFile($file);

        return $simpleExcelWriter;
    }

    public static function createWithoutBom(string $file, string $type = '')
    {
        return static::create(
            file: $file,
            type: $type,
            shouldAddBom: false,
        );
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

    protected function __construct(
        string $path,
        string $type = '',
        ?string $delimiter = null,
        ?bool $shouldAddBom = null,
    ) {
        $this->path = $path;

        $this->csvOptions = new CSVOptions();

        $this->writer = WriterFactory::createFromFile($path);

        if (($delimiter || $shouldAddBom) &&
            $this->writer instanceof Writer) {
            if ($delimiter) {
                $this->csvOptions->FIELD_DELIMITER = $delimiter;
            }

            if ($shouldAddBom) {
                $this->csvOptions->SHOULD_ADD_BOM = $shouldAddBom;
            }

            $this->writer = WriterFactory::createFromFile($path, $this->csvOptions);
        }
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

    public function setHeaderStyle(Style $style)
    {
        $this->headerStyle = $style;

        return $this;
    }

    /**
     * @param \OpenSpout\Common\Entity\Row|array $row
     * @param Style|null $style
     */
    public function addRow($row, Style $style = null)
    {
        if (is_array($row)) {
            if ($this->processHeader && $this->processingFirstRow) {
                $this->writeHeaderFromRow($row);
            }

            $row = Row::fromValues($row, $style);
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
        $headerRow = Row::fromValues($header, $this->headerStyle);

        $this->writer->addRow($headerRow);
        $this->numberOfRows++;

        $this->processingFirstRow = false;

        return $this;
    }

    protected function writeHeaderFromRow(array $row)
    {
        $headerValues = array_keys($row);

        $headerRow = Row::fromValues($headerValues, $this->headerStyle);

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
        if ($this->writer instanceof Writer) {
            $this->csvOptions->FIELD_DELIMITER = $delimiter;
        }

        return $this;
    }

    public function __destruct()
    {
        $this->close();
    }
}
