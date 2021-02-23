<?php

namespace App\Exports;

use App\Car;
use Illuminate\Contracts\View\View;
use Maatwebsite\Excel\Concerns\FromView;
use Maatwebsite\Excel\Concerns\ShouldAutoSize;
use Maatwebsite\Excel\Events\AfterSheet;
use Maatwebsite\Excel\Concerns\WithEvents;

class ReportExport implements FromView, ShouldAutoSize, WithEvents
{
    protected $code = '';
    protected $head = null;
    protected $data = null;

    public function __construct($code, $head, $data)
    {
        $this->code = $code;
        $this->head = $head;
        $this->data = $data;
    }

    // freeze the first row with headings
    public function registerEvents(): array
    {
        return [            
            AfterSheet::class => function(AfterSheet $event) {
                $event->sheet->freezePane('B2');
            },
        ];
    }

    /**
    * @return \Illuminate\Support\Collection
    */
    public function view(): View
    {
        return view('exports.report', [
            'code' => $this->code,
            'head' => $this->head,
            'body' => $this->data
        ]);
    }
}
