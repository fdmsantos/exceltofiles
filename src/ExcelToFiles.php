<?php
namespace ExcelToFiles;

use ReflectionFunction;

class ExcelToFiles
{
	private $template;
	private $excel;
	private $mapping;
	private $alpahbet;
	private $outputdir;
	private $excludeRows;
	private $closures;
	private $filesname;

	public function __construct($options)
	{
		$this->template = file_get_contents($options['template']);
		$this->excel = new \SpreadsheetReader($options['excel']);
		$this->mapping = isSet($options['mapping']) ? $options['mapping'] : [];
		$this->outputdir = isSet($options['outputdir']) ? rtrim($options['outputdir'], '/') . '/' : '';
		$this->excludeRows = isSet($options['excludeRows']) ? $options['excludeRows'] : [] ;
		$this->filesname = isSet($options['filesname']) ? $options['filesname'] : '{A}.txt';
		$this->alpahbet = array_flip(range('A', 'Z'));
		$this->closures = [];
	}

	public function generate()
	{
		foreach ($this->excel as $index => $row) {
			$file = $this->template;
			if (!in_array($index +1, $this->excludeRows)) {
				foreach ($this->mapping as $var => $column) {
					$file = str_replace($var, $row[$this->getIndex($column)], $file);
				}
				foreach ($this->closures as $var => $closure) {
					$file = str_replace($var,$this->executeClosure($row,$closure),$file);
				}
				file_put_contents($this->filename($row) , $file );
			}
		}
	}

	public function mapWithClosure($var, $callback) 
	{
		$this->closures[$var] = $callback;
		return $this;
	}

	private function getIndex($column)
	{
		$columLetters = str_split(strtoupper($column));
		$lastLetter = array_pop($columLetters);
		$index = 0;
		foreach (array_reverse($columLetters) as $item => $letter) {
			$index += pow(count($this->alpahbet),$item + 1) * ($this->alpahbet[$letter] + 1); 
		}
		return $index += $this->alpahbet[$lastLetter];
	}

	private function executeClosure($row,$closure)
	{
		$options = [];
		$reflection = new ReflectionFunction($closure);
 		$lines = file($reflection->getFileName());
		for($l = $reflection->getStartLine(); $l < $reflection->getEndLine(); $l++) {
			if (preg_match_all('/\$columns\[[\'|"](\w+)[\'|"]\]/', $lines[$l],$matches)) {
				foreach($matches[1] as $column) {
					$options[$column] = $row[$this->getIndex(strtoupper($column))];
				}
			}
    	}
    	return $closure($options);
	}

	private function filename($row) 
	{
		$filename = $this->filesname;
		if (preg_match_all('/[^{]*{([A-Z]+)}/mi', $filename,$matches)) {
			foreach($matches[1] as $column) {
				$filename = str_replace('{'.$column.'}',
									$row[$this->getIndex(strtoupper($column))],
									$filename
				);
			}
		}
		return $this->outputdir.$filename;
	}

}