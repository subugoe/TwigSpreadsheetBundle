<?php

namespace MewesK\TwigSpreadsheetBundle\Twig\Node;

use MewesK\TwigSpreadsheetBundle\Wrapper\PhpSpreadsheetWrapper;
use Twig\Node\Node;

abstract class BaseNode extends Node
{
    /**
     * @var string
     */
    public const CODE_FIX_CONTEXT = '$context = '.PhpSpreadsheetWrapper::class.'::fixContext($context);'.\PHP_EOL;

    /**
     * @var string
     */
    public const CODE_INSTANCE = '$context[\''.PhpSpreadsheetWrapper::INSTANCE_KEY."']";

    /**
     * @return string[]
     */
    abstract public function getAllowedParents(): array;
}
