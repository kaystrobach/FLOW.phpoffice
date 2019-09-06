# FLOW.phpoffice

**Attention, there was a an semver violation in the swith from version 1.0.x to 1.1.x**

## Init PHPExcel

```
$loader = new \KayStrobach\PhpOffice\Utility\PHPOfficeUtility('PHPExcel');
$loader->init('PHPExcel');
```

## Same for PHPWord (Just use the classes)

```
$template = new \PhpOffice\PhpWord\TemplateProcessor($this->templatePath);
```

# Documentation

Please look into the appropriate Package description on how to handle the phpOffice libraries: https://github.com/PHPOffice
