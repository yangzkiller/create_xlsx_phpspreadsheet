<?php 

//AUTOLOAD DO COMPOSER
require __DIR__.'/vendor/autoload.php';

//DEPENDENCIAS DO PROJETO
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

//INSTANCIA PRINCIPAL PLANILHA
$spreadsheet = new Spreadsheet();

//OBTÉM A ABA ATIVA DENTRO DO ARQUIVO DO EXCEL
$sheet = $spreadsheet->getActiveSheet();

//DEFINE O CONTEÚDO DA CÉLULA A1 (TITULO DO ARQUIVO)
$sheet->setCellValue('A1', 'CREATE XLSX WITH PHP');

//ESTILOS DA CÉLULA A1
$styles = [
    'font' => [
        'bold'=> true,
        'color' => [
            'rgb' => 'F00F00'
        ],
        'size' => 25,
        'name' => 'Arial'
    ]
];
//DEFINE  O ESTILO DA CÉLULA A1
$sheet->getStyle('A1')->applyFromArray($styles);

//ESTILOS DO CABEÇALHO
$styles = [
    'font' => [
        'bold'=> true,
        'name' => 'Arial'
    ]
];
//DEFINE  O ESTILO NO CABEÇALHO
$sheet->getStyle('A3:C3')->applyFromArray($styles);

/*
//CABEÇALHOS
$sheet->setCellValue('A3', 'ID');
$sheet->setCellValue('B3', 'Nome');
$sheet->setCellValue('C3', 'Valor');

//VALORES PRIMEIRA LINHA
$sheet->setCellValue('A4', '1');
$sheet->setCellValue('B4', 'Monitor LG');
$sheet->setCellValue('C4', '600.00');

//VALORES SEGUNDA LINHA
$sheet->setCellValue('A5', '2');
$sheet->setCellValue('B5', 'Impressora EPSON');
$sheet->setCellValue('C5', '900.00');
*/

//VARIAVEL CONTENDO O ARRAY DE DADOS DA PLANILHA
$cells = [
    ['ID', 'Nome', 'Valor'],
    [1, 'Monitor LG', 600.00],
    [2, 'Impressora EPSON', 900.00],
    [3, 'Notebook HP', 3500.00],
    [null, 'Total', '=SUM(C4:C6)']
];

//DEFINE OS VALORE DENTRO DA PLANILHA UTILIZANDO UM ARRAY
$sheet->fromArray($cells, null, 'A3');

//ESCREVE O ARQUIVO NO DISCO COM O FORMATO XLSX
$writer = new Xlsx($spreadsheet);
$writer->save('arquivo.xlsx');