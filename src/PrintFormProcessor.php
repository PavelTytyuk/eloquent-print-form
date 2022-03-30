<?php

namespace Mnvx\EloquentPrintForm;

use Illuminate\Database\Eloquent\Model;
use Illuminate\Support\Arr;

class PrintFormProcessor
{
    protected string $tmpPrefix;

    protected Pipes $pipes;

    public function __construct(Pipes $pipes = null, string $tmpPrefix = 'pfp_')
    {
        $this->tmpPrefix = $tmpPrefix;
        $this->pipes = $pipes ?? new Pipes();
    }

    /**
     * @param string $templateFile
     * @param Model|array $entity
     * @param callable[] $customSetters Keys are variables, values are callbacks
     * with signature callback(TemplateProcessor $processor, string $variable, ?string $value)
     * @return string Temporary file name with processed document
     * @throws PrintFormException
     */
    public function process(string $templateFile, $entity, array $customSetters = []): string
    {
        try {
            $templateProcessor = new TemplateProcessor($templateFile);
        } catch (\Exception $e) {
            throw new PrintFormException("Cant create processor for '$templateFile'. " . $e->getMessage());
        }
        $tables = new Tables();

        // Process simple fields and collect information about table fields
        foreach ($templateProcessor->getVariables() as $variable) {
            [$current, $isTable] = $this->getValue($entity, $variable, $tables, $templateProcessor);
            if ($isTable) {
                continue;
            }
            if (isset($customSetters[$variable])) {
                $customSetters[$variable]($templateProcessor, $variable, $current);
            }
            else {
                switch (true) {
                    case $current instanceof \PhpOffice\PhpWord\Element\AbstractElement:
                        $templateProcessor->setComplexValue($variable, $current);
                        break;
                    default:
                        $templateProcessor->setValue($variable, htmlspecialchars($current));
                        break;

                }
            }
        }

        // Process table fields
        foreach ($tables->get() as $table) {
            $this->processTable($table, $templateProcessor);
        }
        $tempFolder = resource_path('print_templates/temp');
        $tempFileName = tempnam($tempFolder, $this->tmpPrefix);
        $templateProcessor->saveAs($tempFileName);
        return $tempFileName;
    }

    /**
     * @param Table $table
     * @param TemplateProcessor $templateProcessor
     * @throws PrintFormException
     */
    protected function processTable(Table $table, TemplateProcessor $templateProcessor)
    {
        $marker = $table->marker();
        $values = [];
        $rowNumber = 0;

        foreach ($table->entities() as $entity) {
            $rowNumber++;
            $item = [];
            foreach ($table->variables() as $variable => $shortVariable) {
                $marker = $variable;
                [$item[$variable], $isTable] = $this->getValue($entity, $shortVariable, null, $templateProcessor);
                $item[$variable] = htmlspecialchars($item[$variable]);
            }
            $item[$table->marker() . '#row_number'] = $rowNumber;
            $values[] = $item;
        }
        if (empty($values)) {
            $item = [];
            foreach ($table->variables() as $variable => $shortVariable) {
                $marker = $variable;
                $item[$variable] = $this->pipes->placeholder;
            }
            $item[$table->marker() . '#row_number'] = '_';
            $values[] = $item;
        }
        $templateProcessor->cloneRowAndSetValues($marker, $values);
    }

    /**
     * @param Model|array $entity
     * @param string $variable
     * @param Tables|null $tables
     * @return array [value, is table]
     * @throws PrintFormException
     */
    protected function getValue($entity, string $variable, Tables $tables = null, $templateProcessor): array
    {
        $pipes = explode('|', $variable);
        $parts = explode('.', array_shift($pipes));
        $current = $entity;
        $prefix = '';

        foreach ($parts as $part) {
            $prefix .= $part . '.';
            if (!is_object($current) && ! $current instanceof \Traversable && !is_array($current)) {
                $current = '';
                break;
            }
            try {

                if (is_array($current)) {
                    $current = Arr::get($current, $part);
                } else if (strpos($part, '[') !== false) {
                     preg_match_all('/\[([0-9])+\]/', $part, $matches);
                    $partPrepared = str_replace($matches[0][0], '.'.$matches[1][0], $part);
                    $current = Arr::get($current, $partPrepared);
                } else {
                    $current = $current->$part;
                }
            } catch (\Throwable $e) {
                $current = '';
                break;
            }



            if ($current instanceof \Traversable) {
                if ($tables) {
                    $tables->add($variable, $current, $prefix);
                }
                return [null, true];
            }
        }

        // Process pipes
        foreach ($pipes as $pipe) {
            try {
                $current = $this->pipes->$pipe($current, $templateProcessor, $variable);
            } catch (\Throwable $e) {
                throw $e;
                throw new PrintFormException("Cant process pipe '$pipe' for expression `$variable`. " .
                    $e->getMessage());
            }
        }
        return [$current, false];
   }
}
