<?php

namespace Spatie\SimpleExcel\Tests;

use PHPUnit\Framework\TestCase as PhpUnitTestCase;

abstract class TestCase extends PhpUnitTestCase
{
    public function getStubPath(string $name): string
    {
        return __DIR__."/stubs/{$name}";
    }
}
