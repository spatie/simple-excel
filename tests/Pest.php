<?php

uses(PHPUnit\Framework\TestCase::class)->in('.');

// Functions

function getStubPath(string $name): string
{
    return __DIR__ . "/stubs/{$name}";
}
