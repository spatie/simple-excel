<?php

namespace Spatie\SimpleExcel;

use OpenSpout\Common\Exception\UnsupportedTypeException;

/**
 * @internal This trait is not meant to be used directly.
 */
trait SpreadsheetFactoryTrait
{
    /**
     * @template T of object
     * @param string $type
     * @param array<string, class-string<T>> $map
     */
    protected static function resolveFromType(
        string $type,
        array $map,
        mixed $options = null,
        ?string $message = null
    ): object {
        return new ($map[$type] ?? throw new UnsupportedTypeException(
            $message ?? "Unsupported type: {$type}"
        ))($options);
    }
}
