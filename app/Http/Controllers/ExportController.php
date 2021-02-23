<?php

namespace App\Http\Controllers;

use App\Exports\ReportExport;
use Illuminate\Http\Request;
use Goutte\Client;
use Maatwebsite\Excel\Facades\Excel;

class ExportController extends Controller
{

    public function index() {
        return view('index');
    }

    public function export(Request $request) {
        try{
            $client = new Client();
            $code = $request->code;

            $years = [2010, 2011, 2012, 2013, 2014, 2015, 2016, 2017, 2018, 2019, 2020];
            $data = [
                0 => ["Lợi nhuận sau thuế"],
                1 => ["Tổng cộng tài sản"],
                2 => ["Lợi nhuận sau thuế/tổng tài sản"],
                3 => ["Nợ phải trả"],
                4 => ["Nợ dài hạn"],
                5 => ["Vốn chủ sở hữu"],
                6 => ["Hàng tồn kho"],
                7 => ["Tài sản cố định hữu hình"],
                8 => ["Tài sản cố định"],
                9 => ["Tổng lợi nhuận kế toán trước thuế"], 
                10 => ["Chi phí thuế TNDN hiện hành"],
                11 => ["Lợi nhuận sau thuế thu nhập doanh nghiệp"],
                12 => ["Lợi nhuận sau thuế của công ty mẹ"],
                13 => ["Chi phí lãi vay"]
            ];
            $head = [''];
            foreach ($years as $year) {
                $tab1 = $client->request('GET', 'https://s.cafef.vn/bao-cao-tai-chinh/'.$code.'/BSheet/'.$year.'/4/0/0/ket-qua-hoat-dong-kinh-doanh-cong-ty-co-phan-dau-tu-the-gioi-di-dong.chn');
                $tab2 = $client->request('GET', 'https://s.cafef.vn/bao-cao-tai-chinh/'.$code.'/IncSta/'.$year.'/4/0/0/ket-qua-hoat-dong-kinh-doanh-cong-ty-co-phan-dau-tu-the-gioi-di-dong.chn');

                //head
                $a = $tab1->filter('#tblGridData td:nth-child(2)')->text();
                $b = $tab1->filter('#tblGridData td:nth-child(3)')->text();
                $c = $tab1->filter('#tblGridData td:nth-child(4)')->text();
                $d = $tab1->filter('#tblGridData td:nth-child(5)')->text();
                array_push($head, $a, $b, $c, $d);

                // 0 => ["Lợi nhuận sau thuế"],
                $a = $tab2->filter('#\36 0 > td:nth-child(2)')->text();
                $b = $tab2->filter('#\36 0 > td:nth-child(3)')->text();
                $c = $tab2->filter('#\36 0 > td:nth-child(4)')->text();
                $d = $tab2->filter('#\36 0 > td:nth-child(5)')->text();
                array_push($data[0], $a, $b, $c, $d);

                // 1 => ["Tổng cộng tài sản"],
                $a = $tab1->filter('#\30 01 > td:nth-child(2)')->text();
                $b = $tab1->filter('#\30 01 > td:nth-child(3)')->text();
                $c = $tab1->filter('#\30 01 > td:nth-child(4)')->text();
                $d = $tab1->filter('#\30 01 > td:nth-child(5)')->text();
                array_push($data[1], $a, $b, $c, $d);

                //2 => ["Lợi nhuận sau thuế/tổng tài sản"],

                //3 => ["NỢ PHẢI TRẢ"],
                $a = $tab1->filter('#\33 00 > td:nth-child(2)')->text();
                $b = $tab1->filter('#\33 00 > td:nth-child(3)')->text();
                $c = $tab1->filter('#\33 00 > td:nth-child(4)')->text();
                $d = $tab1->filter('#\33 00 > td:nth-child(5)')->text();
                array_push($data[3], $a, $b, $c, $d);

                //4 => ["Nợ dài hạn"],
                $a = $tab1->filter('#\33 30 > td:nth-child(2)')->text();
                $b = $tab1->filter('#\33 30 > td:nth-child(3)')->text();
                $c = $tab1->filter('#\33 30 > td:nth-child(4)')->text();
                $d = $tab1->filter('#\33 30 > td:nth-child(5)')->text();
                array_push($data[4], $a, $b, $c, $d);


                //5 => ["Vốn chủ sở hữu"],
                $a = $tab1->filter('#\34 10 > td:nth-child(2)')->text();
                $b = $tab1->filter('#\34 10 > td:nth-child(3)')->text();
                $c = $tab1->filter('#\34 10 > td:nth-child(4)')->text();
                $d = $tab1->filter('#\34 10 > td:nth-child(5)')->text();
                array_push($data[5], $a, $b, $c, $d);

                // 6 => ["Hàng tồn kho"],
                $a = $tab1->filter('#\34 10 > td:nth-child(2)')->text();
                $b = $tab1->filter('#\34 10 > td:nth-child(3)')->text();
                $c = $tab1->filter('#\34 10 > td:nth-child(4)')->text();
                $d = $tab1->filter('#\34 10 > td:nth-child(5)')->text();
                array_push($data[6], $a, $b, $c, $d);

                // 7 => ["Tài sản cố định hữu hình"],
                $a = $tab1->filter('#\32 21 > td:nth-child(2)')->text();
                $b = $tab1->filter('#\32 21 > td:nth-child(3)')->text();
                $c = $tab1->filter('#\32 21 > td:nth-child(4)')->text();
                $d = $tab1->filter('#\32 21 > td:nth-child(5)')->text();
                array_push($data[7], $a, $b, $c, $d);

                // 8 => ["Tài sản cố định"],
                $a = $tab1->filter('#\32 20 > td:nth-child(2)')->text();
                $b = $tab1->filter('#\32 20 > td:nth-child(3)')->text();
                $c = $tab1->filter('#\32 20 > td:nth-child(4)')->text();
                $d = $tab1->filter('#\32 20 > td:nth-child(5)')->text();
                array_push($data[8], $a, $b, $c, $d);

                // 9 => ["Tổng lợi nhuận kế toán trước thuế"], 
                $a = $tab2->filter('#\35 0 > td:nth-child(2)')->text();
                $b = $tab2->filter('#\35 0 > td:nth-child(3)')->text();
                $c = $tab2->filter('#\35 0 > td:nth-child(4)')->text();
                $d = $tab2->filter('#\35 0 > td:nth-child(5)')->text();
                array_push($data[9], $a, $b, $c, $d);

                // 10 => ["Chi phí thuế TNDN hiện hành"],
                $a = $tab2->filter('#\35 1 > td:nth-child(2)')->text();
                $b = $tab2->filter('#\35 1 > td:nth-child(3)')->text();
                $c = $tab2->filter('#\35 1 > td:nth-child(4)')->text();
                $d = $tab2->filter('#\35 1 > td:nth-child(5)')->text();
                array_push($data[10], $a, $b, $c, $d);

                // 11 => ["Lợi nhuận sau thuế thu nhập doanh nghiệp"],
                $a = $tab2->filter('#\36 0 > td:nth-child(2)')->text();
                $b = $tab2->filter('#\36 0 > td:nth-child(3)')->text();
                $c = $tab2->filter('#\36 0 > td:nth-child(4)')->text();
                $d = $tab2->filter('#\36 0 > td:nth-child(5)')->text();
                array_push($data[11], $a, $b, $c, $d);

                // 12 => ["Lợi nhuận sau thuế của công ty mẹ"],
                $a = $tab2->filter('#\36 2 > td:nth-child(2)')->text();
                $b = $tab2->filter('#\36 2 > td:nth-child(3)')->text();
                $c = $tab2->filter('#\36 2 > td:nth-child(4)')->text();
                $d = $tab2->filter('#\36 2 > td:nth-child(5)')->text();
                array_push($data[12], $a, $b, $c, $d);

                // 13 => ["Chi phí lãi vay"]
                $a = $tab2->filter('#\32 3 > td:nth-child(2)')->text();
                $b = $tab2->filter('#\32 3 > td:nth-child(3)')->text();
                $c = $tab2->filter('#\32 3 > td:nth-child(4)')->text();
                $d = $tab2->filter('#\32 3 > td:nth-child(5)')->text();
                array_push($data[13], $a, $b, $c, $d);
            }
            //calc 2
            foreach($data[0] as $index => $val) {
                if($index > 0) {
                    $a = str_replace(',', '', $data[0][$index]);
                    $b = str_replace(',', '', $data[1][$index]);
                    $data[2][$index] = $b ? ((float)$a / (float)$b) . '' : 0;
                    $data[2][$index] = str_replace('.', ',', $data[2][$index]);
                }
            }
            // return $data;
            return Excel::download(new ReportExport($code, $head, $data), 'Report_'.$code.'_'.date('YmdHis').'.xlsx');
        }catch(\InvalidArgumentException $e) {
            abort(404);
        }
    }
}
