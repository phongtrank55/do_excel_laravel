<?php

namespace App\ExcelUtilty;

use Excel;
use Maatwebsite\Excel\Concerns\FromCollection;
use Maatwebsite\Excel\Concerns\Exportable;
use Maatwebsite\Excel\Concerns\WithHeadings;
use Maatwebsite\Excel\Concerns\ShouldAutoSize;
use Maatwebsite\Excel\Concerns\WithEvents;
use Maatwebsite\Excel\Events\AfterSheet;

class ExportExcel implements FromCollection, WithHeadings, ShouldAutoSize, WithEvents
{
    use Exportable;

    private $titleIndex; 
    private $needIndex; 
    private $data;
    private $heading;
    private $cellRangeTable;
    private $cellRangeHeading;
    private $cellRangeIndex;

    public function  __construct(array $heading, array $data, $needIndex = true, $titleIndex = 'STT'){
        $this->data = $data;
        $this->heading = $heading;
        $this->needIndex = $needIndex;
        $this->titleIndex = $titleIndex;
        $this->getCellRange();
    }

    public function collection()
    {
        $result = $this->data;

        if ($this->needIndex){
            //thêm cột số thứ tự
            $index = 1;
            foreach ($this->data as $i => $row) {
                $row = (array) $row;
                $result[$i] = array_merge([$index++], $row);
            }
        }
        return (collect($result));
    }

    public function headings(): array{
        return $this->needIndex ? array_merge([$this->titleIndex], $this->heading) : $this->heading;
    }
    public function registerEvents(): array
    {   
        $styleCell = [
            'borders' => [
                'allBorders' => [
                    'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
                    // 'color' => ['argb' => 'FFFF0000'],
                ],
            ],
        ];
        $styleHeading = [
            'font' => [
                'bold' => true,
            ],
            'alignment' => [
                'horizontal' => 'center',
                'vertical' => 'center'
            ]
        ];
        $styleIndex = [
            'alignment' => [
                'horizontal' => 'center',
                'vertical' => 'center'
            ]
        ];
        return [
            AfterSheet::class    => function(AfterSheet $event) use ($styleCell, $styleHeading, $styleIndex) {
                // $event->sheet->getDelegate()->getStyle($cellRange)->getFont()->setSize(14);
                $event->sheet->getDelegate()->getStyle($this->cellRangeTable)->applyFromArray($styleCell);
                $event->sheet->getDelegate()->getStyle($this->cellRangeHeading)->applyFromArray($styleHeading); 
                if($this->needIndex)
                    $event->sheet->getDelegate()->getStyle($this->cellRangeIndex)->applyFromArray($styleIndex); 
            },
        ];
    }
    
    public function download($filename){
        return Excel::download($this, $filename);
    }

    private function getCellRange(){
        $width = $this->needIndex ? count($this->heading) + 1 : count($this->heading); // thêm cột thứ tự
        $heigh = count($this->data) + 1; // Thêm hàng header
        //tính toạ độ
        $alphabet = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z',
        'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO', 'AP', 'AQ', 'AR', 'AS', 'AT', 'AU', 'AV', 'AW', 'AX', 'AY', 'AZ',
        'BA', 'BB', 'BC', 'BD', 'BE', 'BF', 'BG', 'BH', 'BI', 'BJ', 'BK', 'BL', 'BM', 'BN', 'BO', 'BP', 'BQ', 'BR', 'BS', 'BT', 'BU', 'BV', 'BW', 'BX', 'BY', 'BZ',
        'CA', 'CB', 'CC', 'CD', 'CE', 'CF', 'CG', 'CH', 'CI', 'CJ', 'CK', 'CL', 'CM', 'CN', 'CO', 'CP', 'CQ', 'CR', 'CS', 'CT', 'CU', 'CV', 'CW', 'CX', 'CY', 'CZ',
        'DA', 'DB', 'DC', 'DD', 'DE', 'DF', 'DG', 'DH', 'DI', 'DJ', 'DK', 'DL', 'DM', 'DN', 'DO', 'DP', 'DQ', 'DR', 'DS', 'DT', 'DU', 'DV', 'DW', 'DX', 'DY', 'DZ',
        'EA', 'EB', 'EC', 'ED', 'EE', 'EF', 'EG', 'EH', 'EI', 'EJ', 'EK', 'EL', 'EM', 'EN', 'EO', 'EP', 'EQ', 'ER', 'ES', 'ET', 'EU', 'EV', 'EW', 'EX', 'EY', 'EZ'];

        $this->cellRangeTable="A1:" . $alphabet[$width-1] . $heigh; // bảng
        $this->cellRangeHeading = "A1:" . $alphabet[$width-1] . '1'; // tiêu đề
        if($this->needIndex)
            $this->cellRangeIndex = "A2:A" . $heigh; // cột số thứ tự
    }
}