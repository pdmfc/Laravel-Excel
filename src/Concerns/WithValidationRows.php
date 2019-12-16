<?php

namespace Maatwebsite\Excel\Concerns;

interface WithValidationRows
{
    /**
     * @param array $rows
     * 
     * @return array
     */
    public function rules(array $rows): array;
}
