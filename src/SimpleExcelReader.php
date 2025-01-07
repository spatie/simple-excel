<?php

namespace Spatie\SimpleExcel;

use Illuminate\Support\LazyCollection;
use InvalidArgumentException;
use OpenSpout\Common\Entity\Cell;
use OpenSpout\Common\Entity\Cell\FormulaCell;
use OpenSpout\Common\Entity\Row;
use OpenSpout\Reader\CSV\Options as CSVOptions;
use OpenSpout\Reader\CSV\Reader as CSVReader;
use OpenSpout\Reader\ReaderInterface;
use OpenSpout\Reader\RowIteratorInterface;
use OpenSpout\Reader\SheetInterface;
use OpenSpout\Reader\XLSX\Options as XLSXOptions;
use OpenSpout\Reader\XLSX\Reader as XLSXReader;

class SimpleExcelReader
{
    protected ReaderInterface $reader;
    protected RowIteratorInterface $rowIterator;
    protected int $sheetNumber = 1;
    protected string $sheetName = "";
    protected bool $searchSheetByName = false;
    protected bool $processHeader = true;
    protected bool $trimHeader = false;
    protected bool $headersToSnakeCase = false;
    protected bool $parseFormulas = true;
    protected ?string $trimHeaderCharacters = null;
    protected mixed $formatHeadersUsing = null;
    protected ?array $headers = null;
    protected int $headerOnRow = 0;
    protected ?array $customHeaders = [];
    protected int $skip = 0;
    protected int $limit = 0;
    protected bool $useLimit = false;
    protected CSVOptions $csvOptions;
    protected XLSXOptions $xlsxOptions;

    public static function create(string $file): static
    {
        return new static($file);
    }

    public function __construct(protected string $path)
    {
        $this->csvOptions = new CSVOptions();
        $this->xlsxOptions = new XLSXOptions();

        $this->reader = ReaderFactory::createFromFile($this->path);

        $this->setReader();
    }

    protected function setReader(): void
    {
        $options = match (true) {
            $this->reader instanceof CSVReader => $this->csvOptions,
            $this->reader instanceof XLSXReader => $this->xlsxOptions,
            default => null,
        };

        $this->reader = ReaderFactory::createFromFile($this->path, $options);
    }

    public function getPath(): string
    {
        return $this->path;
    }

    public function headerOnRow(int $headerRow): self
    {
        $this->headerOnRow = $headerRow;

        return $this;
    }

    public function noHeaderRow(): self
    {
        $this->processHeader = false;

        return $this;
    }

    public function useHeaders(array $headers): self
    {
        $this->customHeaders = $headers;

        return $this;
    }

    public function useDelimiter(string $delimiter): self
    {
        if ($this->reader instanceof CSVReader) {
            $this->csvOptions->FIELD_DELIMITER = $delimiter;
        }

        return $this;
    }

    public function useFieldEnclosure(string $fieldEnclosure): self
    {
        if ($this->reader instanceof CSVReader) {
            $this->csvOptions->FIELD_ENCLOSURE = $fieldEnclosure;
        }

        return $this;
    }

    public function useEncoding(string $encoding): self
    {
        if ($this->reader instanceof CSVReader) {
            $this->csvOptions->ENCODING = $encoding;
        }

        return $this;
    }

    public function preserveDateTimeFormatting(): self
    {
        if ($this->reader instanceof XLSXReader) {
            $this->xlsxOptions->SHOULD_FORMAT_DATES = true;
        }

        return $this;
    }

    public function preserveEmptyRows(): self
    {
        if ($this->reader instanceof XLSXReader) {
            $this->xlsxOptions->SHOULD_PRESERVE_EMPTY_ROWS = true;
        }

        return $this;
    }

    public function trimHeaderRow(?string $characters = null): self
    {
        $this->trimHeader = true;
        $this->trimHeaderCharacters = $characters;

        return $this;
    }

    public function formatHeadersUsing(callable $callback): self
    {
        $this->formatHeadersUsing = $callback;

        return $this;
    }

    public function headersToSnakeCase(): self
    {
        $this->headersToSnakeCase = true;

        return $this;
    }

    public function keepFormulas()
    {
        $this->parseFormulas = false;

        return $this;
    }

    public function getReader(): ReaderInterface
    {
        return $this->reader;
    }

    public function skip(int $count): SimpleExcelReader
    {
        $this->skip = $count;

        return $this;
    }

    public function take(int $count): SimpleExcelReader
    {
        $this->limit = $count;
        $this->useLimit = true;

        return $this;
    }

    public function hasSheet(string $sheetName): bool
    {
        $this->setReader();

        $this->reader->open($this->path);

        foreach ($this->reader->getSheetIterator() as $sheet) {
            if ($sheetName != "" && $sheetName === $sheet->getName()) {
                return true;
            }
        }

        return false;
    }

    public function fromSheet(int $sheetNumber): SimpleExcelReader
    {
        $this->sheetNumber = $sheetNumber;
        $this->headers = null;
        $this->searchSheetByName = false;

        return $this;
    }

    public function fromSheetName(string $sheetName): SimpleExcelReader
    {
        $this->searchSheetByName = true;
        $this->sheetName = $sheetName;
        $this->headers = null;

        return $this;
    }

    public function getRows(): LazyCollection
    {
        $sheet = $this->getSheet();

        $this->rowIterator = $sheet->getRowIterator();

        $this->rowIterator->rewind();

        if ($this->processHeader) {
            $this->getHeaders();
            $this->rowIterator->next();
        }

        return LazyCollection::make(function () {
            while ($this->rowIterator->valid() && $this->skip && $this->skip--) {
                $this->rowIterator->next();
            }
            while ($this->rowIterator->valid() && (! $this->useLimit || $this->limit--)) {
                $row = $this->rowIterator->current();

                yield $this->getValueFromRow($row);

                $this->rowIterator->next();
            }
        });
    }

    public function getHeaders(): ?array
    {
        if (! $this->processHeader) {
            if ($this->customHeaders) {
                return $this->customHeaders;
            }

            return null;
        }

        if ($this->headers) {
            return $this->headers;
        }

        $sheet = $this->getSheet();

        $this->rowIterator = $sheet->getRowIterator();

        $this->rowIterator->rewind();

        $headerRow = $this->rowIterator->current();

        if ($this->headerOnRow > 0) {
            $skip = $this->headerOnRow;
            while ($skip--) {
                $this->rowIterator->next();
            }
            $headerRow = $this->rowIterator->current();
        }

        if (is_null($headerRow)) {
            $this->noHeaderRow();

            return null;
        }

        $this->headers = $this->processHeaderRow($headerRow->toArray());

        if ($this->customHeaders) {
            return $this->customHeaders;
        }

        return $this->headers;
    }

    public function getOriginalHeaders(): ?array
    {
        $this->getHeaders();

        return $this->headers;
    }

    protected function processHeaderRow(array $headers): array
    {
        if ($this->trimHeader) {
            $headers = $this->convertHeaders([$this, 'trim'], $headers);
        }

        if ($this->headersToSnakeCase) {
            $headers = $this->convertHeaders([$this, 'toSnakecase'], $headers);
        }

        if ($this->formatHeadersUsing) {
            $headers = $this->convertHeaders($this->formatHeadersUsing, $headers);
        }

        return $headers;
    }

    protected function convertHeaders(callable $callback, array $headers): array
    {
        return array_map(fn ($header) => $callback($header), $headers);
    }

    protected function trim(string $header): string
    {
        $arguments = [];
        $arguments[] = $header;

        if (! is_null($this->trimHeaderCharacters)) {
            $arguments[] = $this->trimHeaderCharacters;
        }

        return trim(...$arguments);
    }

    protected function toSnakeCase(string $header): string
    {
        return str_replace(
            ' ',
            '_',
            strtolower(preg_replace('/(?<=\d)(?=[A-Za-z])|(?<=[A-Za-z])(?=\d)|(?<=[a-z])(?=[A-Z])/', '_', trim($header)))
        );
    }

    protected function getValueFromRow(Row $row): array
    {
        $values = array_map(function (Cell $cell) {
            return $cell instanceof FormulaCell && $this->parseFormulas
                ? $cell->getComputedValue()
                : $cell->getValue();
        }, $row->getCells());

        ksort($values);

        $headers = $this->customHeaders ?: $this->headers;

        if (! $headers) {
            return $values;
        }

        $values = array_slice($values, 0, count($headers));

        while (count($values) < count($headers)) {
            $values[] = '';
        }

        return array_combine($headers, $values);
    }

    protected function getSheet(): SheetInterface
    {
        $this->setReader();

        $this->reader->open($this->path);

        return ($this->searchSheetByName) ? $this->getActiveSheetByName() : $this->getActiveSheetByIndex();
    }

    protected function getActiveSheetByName(): SheetInterface
    {
        $sheet = null;
        foreach ($this->reader->getSheetIterator() as $key => $sheet) {
            if ($this->sheetName != "" && $this->sheetName === $sheet->getName()) {
                break;
            }
        }
        if ($this->sheetName != "" && $this->sheetName !== $sheet->getName()) {
            throw new InvalidArgumentException("Sheet name {$this->sheetName} does not exist in {$this->path}.");
        }

        return $sheet;
    }

    protected function getActiveSheetByIndex(): SheetInterface
    {
        $key = null;
        $sheet = null;
        foreach ($this->reader->getSheetIterator() as $key => $sheet) {
            if ($key === $this->sheetNumber) {
                break;
            }
        }
        if ($this->sheetNumber !== $key) {
            throw new InvalidArgumentException("Sheet Index {$key} does not exist in {$this->path}.");
        }

        return $sheet;
    }

    public function close(): void
    {
        $this->reader->close();
    }

    public function __destruct()
    {
        $this->close();
    }
}
