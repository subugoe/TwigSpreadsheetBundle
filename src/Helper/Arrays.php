<?php

namespace MewesK\TwigSpreadsheetBundle\Helper;

class Arrays
{
    /**
     * Merge two arrays recursively. Works the same as 'array_merge_recursive' but will only merge array values, other
     * values are overridden.
     */
    public static function mergeRecursive(array &$array1, array &$array2): array
    {
        foreach ($array2 as $key => &$value) {
            $array1[$key] = \is_array($value) && isset($array1[$key]) && \is_array($array1[$key]) ?
                self::mergeRecursive($array1[$key], $value) :
                $value;
        }

        return $array1;
    }
}
