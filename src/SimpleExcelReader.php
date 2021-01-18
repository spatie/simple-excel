<?php

namespace Spatie\SimpleExcel;

use Box\Spout\Common\Entity\Row;
use Box\Spout\Reader\Common\Creator\ReaderEntityFactory;
use Box\Spout\Reader\Common\Creator\ReaderFactory;
use Box\Spout\Reader\IteratorInterface;
use Box\Spout\Reader\ReaderInterface;
use Closure;
use Illuminate\Support\LazyCollection;

class SimpleExcelReader
{
    protected string $path;

    protected string $type;

    protected ReaderInterface $reader;

    protected IteratorInterface $rowIterator;

    protected bool $processHeader = true;

    protected bool $trimHeader = false;

    protected bool $headersToSnakeCase = false;

    protected ?string $trimHeaderCharacters = null;

    protected ?Closure $formatHeadersUsing = null;

    protected ?array $headers = null;

    protected int $skip = 0;

    protected int $limit = 0;

    protected bool $useLimit = false;

    public static function create(string $file, string $type = '')
    {
        return new static($file, $type);
    }

    public function __construct(string $path, string $type = '')
    {
        $this->path = $path;

        $this->type = $type;

        $this->reader = $type ?
            ReaderFactory::createFromType($type) :
            ReaderEntityFactory::createReaderFromFile($this->path);
    }

    public function getPath(): string
    {
        return $this->path;
    }

    public function noHeaderRow(): self
    {
        $this->processHeader = false;

        return $this;
    }

    public function useDelimiter(string $delimiter): self
    {
        $this->reader->setFieldDelimiter($delimiter);

        return $this;
    }

    public function useFieldEnclosure(string $fieldEnclosure): self
    {
        $this->reader->setFieldEnclosure($fieldEnclosure);

        return $this;
    }

    public function trimHeaderRow(string $characters = null): self
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

    public function getRows(): LazyCollection
    {
        $this->reader->open($this->path);

        $this->reader->getSheetIterator()->rewind();

        $sheet = $this->reader->getSheetIterator()->current();

        $this->rowIterator = $sheet->getRowIterator();

        $this->rowIterator->rewind();

        /** @var \Box\Spout\Common\Entity\Row $firstRow */
        $firstRow = $this->rowIterator->current();

        if (is_null($firstRow)) {
            $this->noHeaderRow();
        }

        if ($this->processHeader) {
            $this->headers = $this->processHeaderRow($firstRow->toArray());

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
            return null;
        }

        if ($this->headers) {
            return $this->headers;
        }

        $reader = $this->type ?
            ReaderFactory::createFromType($this->type) :
            ReaderEntityFactory::createReaderFromFile($this->path);

        $reader->open($this->path);

        $reader->getSheetIterator()->rewind();

        $sheet = $reader->getSheetIterator()->current();

        $this->rowIterator = $sheet->getRowIterator();

        $this->rowIterator->rewind();

        /** @var \Box\Spout\Common\Entity\Row $firstRow */
        $firstRow = $this->rowIterator->current();

        if (is_null($firstRow)) {
            $this->noHeaderRow();

            return null;
        }

        $this->headers = $this->processHeaderRow($firstRow->toArray());

        return $this->headers;
    }

    public function close()
    {
        $this->reader->close();
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
        return array_map(function ($header) use ($callback) {
            return call_user_func($callback, $header);
        }, $headers);
    }

    public function headerRowFormatter(callable $callback)
    {
        $this->headerRowFormatter = $callback;

        return $this;
    }

    protected function trim(string $header): string
    {
        return call_user_func_array('trim', array_filter([$header, $this->trimHeaderCharacters]));
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
        $values = $row->toArray();
        ksort($values);

        if (! $this->processHeader) {
            return $values;
        }

        $values = array_slice($values, 0, count($this->headers));

        while (count($values) < count($this->headers)) {
            $values[] = '';
        }

        return array_combine($this->headers, $values);
    }

    public function __destruct()
    {
        $this->close();
    }
}
