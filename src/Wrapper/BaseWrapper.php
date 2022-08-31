<?php

namespace MewesK\TwigSpreadsheetBundle\Wrapper;

abstract class BaseWrapper
{
    protected array $parameters = [];

    protected array $mappings = [];

    public function __construct(protected array $context, protected \Twig\Environment $environment)
    {
        $this->mappings = $this->configureMappings();
    }

    public function getParameters(): array
    {
        return $this->parameters;
    }

    public function setParameters(array $parameters)
    {
        $this->parameters = $parameters;
    }

    public function getMappings(): array
    {
        return $this->mappings;
    }

    public function setMappings(array $mappings): void
    {
        $this->mappings = $mappings;
    }

    protected function configureMappings(): array
    {
        return [];
    }

    /**
     * Calls the matching mapping callable for each property.
     *
     * @throws \RuntimeException
     */
    protected function setProperties(array $properties, array $mappings = null, string $column = null): void
    {
        if (null === $mappings) {
            $mappings = $this->mappings;
        }

        foreach ($properties as $key => $value) {
            if (!isset($mappings[$key])) {
                throw new \RuntimeException(sprintf('Missing mapping for key "%s"', $key));
            }

            if (\is_array($value) && \is_array($mappings[$key])) {
                // recursion
                if (isset($mappings[$key]['__multi'])) {
                    // handle multi target structure (with columns)
                    foreach ($value as $_column => $_value) {
                        $this->setProperties($_value, $mappings[$key], $_column);
                    }
                } else {
                    // handle single target structure
                    $this->setProperties($value, $mappings[$key]);
                }
            } elseif (\is_callable($mappings[$key])) {
                // call single and multi target mapping
                // if column is set it is used to get object from the callback in __multi
                $mappings[$key](
                    $value,
                    null !== $column ? $mappings['__multi']($column) : null
                );
            } else {
                throw new \RuntimeException(sprintf('Invalid mapping for key "%s"', $key));
            }
        }
    }
}
