<?php

namespace App\Http\Controllers;

use App\Exports\ReportExport;
use Illuminate\Http\Request;
use Goutte\Client;
use Illuminate\Support\Facades\Storage;
use Maatwebsite\Excel\Facades\Excel;
use ZipArchive;

class ExportController extends Controller
{
    public function __construct(){
        ini_set('max_execution_time', 0);
    }


    public function index() {
        return view('index');
    }
    public function all() {
        return view('all');
    }

    public function export(Request $request) {
        $years = [];
        for($i = $request->year_from; $i<=$request->year_to; $i++) {
            array_push($years, $i);
        }
        if($request->horizontal) {
            return $this->exportHorizontal($request, $years);
        }
        return $this->exportVertical($request, $years);
    }

    public function exportVertical(Request $request, $years) {
        try{
            $head = [
                "",
                "Lợi nhuận sau thuế",
                "Tổng cộng tài sản",
                "Lợi nhuận sau thuế/tổng tài sản",
                "Nợ phải trả",
                "Nợ dài hạn",
                "Vốn chủ sở hữu",
                "Hàng tồn kho",
                "Tài sản cố định hữu hình",
                "Tài sản cố định",
                "Tổng lợi nhuận kế toán trước thuế",
                "Chi phí thuế TNDN hiện hành",
                "Lợi nhuận sau thuế thu nhập doanh nghiệp",
                "Lợi nhuận sau thuế của công ty mẹ",
                "Chi phí lãi vay"
            ];
            $client = new Client();
            $code = $request->code;

            $data = [];
            foreach ($years as $key => $year) {
                $tab1 = $client->request('GET', 'https://s.cafef.vn/bao-cao-tai-chinh/'.$code.'/BSheet/'.$year.'/4/0/0/ket-qua-hoat-dong-kinh-doanh-cong-ty-co-phan-dau-tu-the-gioi-di-dong.chn');
                $tab2 = $client->request('GET', 'https://s.cafef.vn/bao-cao-tai-chinh/'.$code.'/IncSta/'.$year.'/4/0/0/ket-qua-hoat-dong-kinh-doanh-cong-ty-co-phan-dau-tu-the-gioi-di-dong.chn');

                //quy 1
                $a = str_replace(',', '', $tab1->filter('#tblGridData td:nth-child(2)')->text());
                $a0 = $tab1->filter('#\34 20 > td:nth-child(2)')->count() > 0 ? str_replace(',', '', $tab1->filter('#\34 20 > td:nth-child(2)')->text()) : '';
                $a1 = $tab1->filter('#\30 01 > td:nth-child(2)')->count() > 0 ? str_replace(',', '', $tab1->filter('#\30 01 > td:nth-child(2)')->text()) : '';
                $a2 = is_numeric($a0) && is_numeric($a1) && $a1!=0 ? $a0/$a1 : 0;
                $a3 = $tab1->filter('#\33 00 > td:nth-child(2)')->count() > 0 ? str_replace(',', '', $tab1->filter('#\33 00 > td:nth-child(2)')->text()) : '';
                $a4 = $tab1->filter('#\33 30 > td:nth-child(2)')->count() > 0 ? str_replace(',', '', $tab1->filter('#\33 30 > td:nth-child(2)')->text()) : '';
                $a5 = $tab1->filter('#\34 10 > td:nth-child(2)')->count() > 0 ? str_replace(',', '', $tab1->filter('#\34 10 > td:nth-child(2)')->text()) : '';
                $a6 = $tab1->filter('#\31 40 > td:nth-child(2)')->count() > 0 ? str_replace(',', '', $tab1->filter('#\31 40 > td:nth-child(2)')->text()) : '';
                $a7 = $tab1->filter('#\32 21 > td:nth-child(2)')->count() > 0 ? str_replace(',', '', $tab1->filter('#\32 21 > td:nth-child(2)')->text()) : '';
                $a8 = $tab1->filter('#\32 20 > td:nth-child(2)')->count() > 0 ? str_replace(',', '', $tab1->filter('#\32 20 > td:nth-child(2)')->text()) : '';
                $a9 = $tab2->filter('#\35 0 > td:nth-child(2)')->count() > 0 ? str_replace(',', '', $tab2->filter('#\35 0 > td:nth-child(2)')->text()) : '';
                $a10 = $tab2->filter('#\35 1 > td:nth-child(2)')->count() > 0 ? str_replace(',', '', $tab2->filter('#\35 1 > td:nth-child(2)')->text()) : '';
                $a11 = $tab2->filter('#\36 0 > td:nth-child(2)')->count() > 0 ? str_replace(',', '', $tab2->filter('#\36 0 > td:nth-child(2)')->text()) : '';
                $a12 = $tab2->filter('#\36 2 > td:nth-child(2)')->count() > 0 ? str_replace(',', '', $tab2->filter('#\36 2 > td:nth-child(2)')->text()) : '';
                $a13 = $tab2->filter('#\32 3 > td:nth-child(2)')->count() > 0 ? str_replace(',', '', $tab2->filter('#\32 3 > td:nth-child(2)')->text()) : '';
                array_push($data, [$a, $a0, $a1, $a2, $a3, $a4, $a5, $a6, $a7, $a8, $a9, $a10, $a11, $a12, $a13]);

                //quy 2
                $a = str_replace(',', '', $tab1->filter('#tblGridData td:nth-child(3)')->text());
                $a0 = $tab1->filter('#\34 20 > td:nth-child(3)')->count() > 0 ? str_replace(',', '', $tab1->filter('#\34 20 > td:nth-child(3)')->text()) : '';
                $a1 = $tab1->filter('#\30 01 > td:nth-child(3)')->count() > 0 ? str_replace(',', '', $tab1->filter('#\30 01 > td:nth-child(3)')->text()) : '';
                $a2 = is_numeric($a0) && is_numeric($a1) && $a1!=0 ? $a0/$a1 : 0;
                $a3 = $tab1->filter('#\33 00 > td:nth-child(3)')->count() > 0 ? str_replace(',', '', $tab1->filter('#\33 00 > td:nth-child(3)')->text()) : '';
                $a4 = $tab1->filter('#\33 30 > td:nth-child(3)')->count() > 0 ? str_replace(',', '', $tab1->filter('#\33 30 > td:nth-child(3)')->text()) : '';
                $a5 = $tab1->filter('#\34 10 > td:nth-child(3)')->count() > 0 ? str_replace(',', '', $tab1->filter('#\34 10 > td:nth-child(3)')->text()) : '';
                $a6 = $tab1->filter('#\31 40 > td:nth-child(3)')->count() > 0 ? str_replace(',', '', $tab1->filter('#\31 40 > td:nth-child(3)')->text()) : '';
                $a7 = $tab1->filter('#\32 21 > td:nth-child(3)')->count() > 0 ? str_replace(',', '', $tab1->filter('#\32 21 > td:nth-child(3)')->text()) : '';
                $a8 = $tab1->filter('#\32 20 > td:nth-child(3)')->count() > 0 ? str_replace(',', '', $tab1->filter('#\32 20 > td:nth-child(3)')->text()) : '';
                $a9 = $tab2->filter('#\35 0 > td:nth-child(3)')->count() > 0 ? str_replace(',', '', $tab2->filter('#\35 0 > td:nth-child(3)')->text()) : '';
                $a10 = $tab2->filter('#\35 1 > td:nth-child(3)')->count() > 0 ? str_replace(',', '', $tab2->filter('#\35 1 > td:nth-child(3)')->text()) : '';
                $a11 = $tab2->filter('#\36 0 > td:nth-child(3)')->count() > 0 ? str_replace(',', '', $tab2->filter('#\36 0 > td:nth-child(3)')->text()) : '';
                $a12 = $tab2->filter('#\36 2 > td:nth-child(3)')->count() > 0 ? str_replace(',', '', $tab2->filter('#\36 2 > td:nth-child(3)')->text()) : '';
                $a13 = $tab2->filter('#\32 3 > td:nth-child(3)')->count() > 0 ? str_replace(',', '', $tab2->filter('#\32 3 > td:nth-child(3)')->text()) : '';
                array_push($data, [$a, $a0, $a1, $a2, $a3, $a4, $a5, $a6, $a7, $a8, $a9, $a10, $a11, $a12, $a13]);

                //quy 3
                $a = str_replace(',', '', $tab1->filter('#tblGridData td:nth-child(4)')->text());
                $a0 = $tab1->filter('#\34 20 > td:nth-child(4)')->count() > 0 ? str_replace(',', '', $tab1->filter('#\34 20 > td:nth-child(4)')->text()) : '';
                $a1 = $tab1->filter('#\30 01 > td:nth-child(4)')->count() > 0 ? str_replace(',', '', $tab1->filter('#\30 01 > td:nth-child(4)')->text()) : '';
                $a2 = is_numeric($a0) && is_numeric($a1) && $a1!=0 ? $a0/$a1 : 0;
                $a3 = $tab1->filter('#\33 00 > td:nth-child(4)')->count() > 0 ? str_replace(',', '', $tab1->filter('#\33 00 > td:nth-child(4)')->text()) : '';
                $a4 = $tab1->filter('#\33 30 > td:nth-child(4)')->count() > 0 ? str_replace(',', '', $tab1->filter('#\33 30 > td:nth-child(4)')->text()) : '';
                $a5 = $tab1->filter('#\34 10 > td:nth-child(4)')->count() > 0 ? str_replace(',', '', $tab1->filter('#\34 10 > td:nth-child(4)')->text()) : '';
                $a6 = $tab1->filter('#\31 40 > td:nth-child(4)')->count() > 0 ? str_replace(',', '', $tab1->filter('#\31 40 > td:nth-child(4)')->text()) : '';
                $a7 = $tab1->filter('#\32 21 > td:nth-child(4)')->count() > 0 ? str_replace(',', '', $tab1->filter('#\32 21 > td:nth-child(4)')->text()) : '';
                $a8 = $tab1->filter('#\32 20 > td:nth-child(4)')->count() > 0 ? str_replace(',', '', $tab1->filter('#\32 20 > td:nth-child(4)')->text()) : '';
                $a9 = $tab2->filter('#\35 0 > td:nth-child(4)')->count() > 0 ? str_replace(',', '', $tab2->filter('#\35 0 > td:nth-child(4)')->text()) : '';
                $a10 = $tab2->filter('#\35 1 > td:nth-child(4)')->count() > 0 ? str_replace(',', '', $tab2->filter('#\35 1 > td:nth-child(4)')->text()) : '';
                $a11 = $tab2->filter('#\36 0 > td:nth-child(4)')->count() > 0 ? str_replace(',', '', $tab2->filter('#\36 0 > td:nth-child(4)')->text()) : '';
                $a12 = $tab2->filter('#\36 2 > td:nth-child(4)')->count() > 0 ? str_replace(',', '', $tab2->filter('#\36 2 > td:nth-child(4)')->text()) : '';
                $a13 = $tab2->filter('#\32 3 > td:nth-child(4)')->count() > 0 ? str_replace(',', '', $tab2->filter('#\32 3 > td:nth-child(4)')->text()) : '';
                array_push($data, [$a, $a0, $a1, $a2, $a3, $a4, $a5, $a6, $a7, $a8, $a9, $a10, $a11, $a12, $a13]);

                //quy 4
                $a = str_replace(',', '', $tab1->filter('#tblGridData td:nth-child(5)')->text());
                $a0 = $tab1->filter('#\34 20 > td:nth-child(5)')->count() > 0 ? str_replace(',', '', $tab1->filter('#\34 20 > td:nth-child(5)')->text()) : '';
                $a1 = $tab1->filter('#\30 01 > td:nth-child(5)')->count() > 0 ? str_replace(',', '', $tab1->filter('#\30 01 > td:nth-child(5)')->text()) : '';
                $a2 = is_numeric($a0) && is_numeric($a1) && $a1!=0 ? $a0/$a1 : 0;
                $a3 = $tab1->filter('#\33 00 > td:nth-child(5)')->count() > 0 ? str_replace(',', '', $tab1->filter('#\33 00 > td:nth-child(5)')->text()) : '';
                $a4 = $tab1->filter('#\33 30 > td:nth-child(5)')->count() > 0 ? str_replace(',', '', $tab1->filter('#\33 30 > td:nth-child(5)')->text()) : '';
                $a5 = $tab1->filter('#\34 10 > td:nth-child(5)')->count() > 0 ? str_replace(',', '', $tab1->filter('#\34 10 > td:nth-child(5)')->text()) : '';
                $a6 = $tab1->filter('#\31 40 > td:nth-child(5)')->count() > 0 ? str_replace(',', '', $tab1->filter('#\31 40 > td:nth-child(5)')->text()) : '';
                $a7 = $tab1->filter('#\32 21 > td:nth-child(5)')->count() > 0 ? str_replace(',', '', $tab1->filter('#\32 21 > td:nth-child(5)')->text()) : '';
                $a8 = $tab1->filter('#\32 20 > td:nth-child(5)')->count() > 0 ? str_replace(',', '', $tab1->filter('#\32 20 > td:nth-child(5)')->text()) : '';
                $a9 = $tab2->filter('#\35 0 > td:nth-child(5)')->count() > 0 ? str_replace(',', '', $tab2->filter('#\35 0 > td:nth-child(5)')->text()) : '';
                $a10 = $tab2->filter('#\35 1 > td:nth-child(5)')->count() > 0 ? str_replace(',', '', $tab2->filter('#\35 1 > td:nth-child(5)')->text()) : '';
                $a11 = $tab2->filter('#\36 0 > td:nth-child(5)')->count() > 0 ? str_replace(',', '', $tab2->filter('#\36 0 > td:nth-child(5)')->text()) : '';
                $a12 = $tab2->filter('#\36 2 > td:nth-child(5)')->count() > 0 ? str_replace(',', '', $tab2->filter('#\36 2 > td:nth-child(5)')->text()) : '';
                $a13 = $tab2->filter('#\32 3 > td:nth-child(5)')->count() > 0 ? str_replace(',', '', $tab2->filter('#\32 3 > td:nth-child(5)')->text()) : '';
                array_push($data, [$a, $a0, $a1, $a2, $a3, $a4, $a5, $a6, $a7, $a8, $a9, $a10, $a11, $a12, $a13]);
            }
            return Excel::download(new ReportExport($code, $head, $data), 'Report_'.$code.'_Vertical_'.date('YmdHis').'.xlsx');
        }catch(\InvalidArgumentException $e) {
            abort(404);
        }
    }

    public function exportHorizontal(Request $request, $years) {
        try{
            $client = new Client();
            $code = $request->code;

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
                $a = str_replace(',', '', $tab1->filter('#\34 20 > td:nth-child(2)')->text());
                $b = str_replace(',', '', $tab1->filter('#\34 20 > td:nth-child(3)')->text());
                $c = str_replace(',', '', $tab1->filter('#\34 20 > td:nth-child(4)')->text());
                $d = str_replace(',', '', $tab1->filter('#\34 20 > td:nth-child(5)')->text());
                array_push($data[0], $a, $b, $c, $d);
                

                // 1 => ["Tổng cộng tài sản"],
                $a = str_replace(',', '', $tab1->filter('#\30 01 > td:nth-child(2)')->text());
                $b = str_replace(',', '', $tab1->filter('#\30 01 > td:nth-child(3)')->text());
                $c = str_replace(',', '', $tab1->filter('#\30 01 > td:nth-child(4)')->text());
                $d = str_replace(',', '', $tab1->filter('#\30 01 > td:nth-child(5)')->text());
                array_push($data[1], $a, $b, $c, $d);
                

                //2 => ["Lợi nhuận sau thuế/tổng tài sản"],

                //3 => ["NỢ PHẢI TRẢ"],
                $a = str_replace(',', '', $tab1->filter('#\33 00 > td:nth-child(2)')->text());
                $b = str_replace(',', '', $tab1->filter('#\33 00 > td:nth-child(3)')->text());
                $c = str_replace(',', '', $tab1->filter('#\33 00 > td:nth-child(4)')->text());
                $d = str_replace(',', '', $tab1->filter('#\33 00 > td:nth-child(5)')->text());
                array_push($data[3], $a, $b, $c, $d);
                

                //4 => ["Nợ dài hạn"],
                if($tab1->filter('#\33 30 > td:nth-child(2)')->count() > 0) {
                    $a = str_replace(',', '', $tab1->filter('#\33 30 > td:nth-child(2)')->text());
                    $b = str_replace(',', '', $tab1->filter('#\33 30 > td:nth-child(3)')->text());
                    $c = str_replace(',', '', $tab1->filter('#\33 30 > td:nth-child(4)')->text());
                    $d = str_replace(',', '', $tab1->filter('#\33 30 > td:nth-child(5)')->text());
                    array_push($data[4], $a, $b, $c, $d);
                }


                //5 => ["Vốn chủ sở hữu"],
                // if(->count() > 0) {}
                if($tab1->filter('#\34 00 > td:nth-child(2)')->count() > 0) {
                    $a = str_replace(',', '', $tab1->filter('#\34 00 > td:nth-child(2)')->text());
                    $b = str_replace(',', '', $tab1->filter('#\34 10 > td:nth-child(3)')->text());
                    $c = str_replace(',', '', $tab1->filter('#\34 10 > td:nth-child(4)')->text());
                    $d = str_replace(',', '', $tab1->filter('#\34 10 > td:nth-child(5)')->text());
                    array_push($data[5], $a, $b, $c, $d);
                }
                

                // // 6 => ["Hàng tồn kho"],
                if($tab1->filter('#\31 40 > td:nth-child(2)')->count() > 0) {
                    $a = str_replace(',', '', $tab1->filter('#\31 40 > td:nth-child(2)')->text());
                    $b = str_replace(',', '', $tab1->filter('#\31 40 > td:nth-child(3)')->text());
                    $c = str_replace(',', '', $tab1->filter('#\31 40 > td:nth-child(4)')->text());
                    $d = str_replace(',', '', $tab1->filter('#\31 40 > td:nth-child(5)')->text());
                    array_push($data[6], $a, $b, $c, $d);
                }

                // // 7 => ["Tài sản cố định hữu hình"],
                if($tab1->filter('#\32 21 > td:nth-child(2)')->count() > 0) {
                    $a = str_replace(',', '', $tab1->filter('#\32 21 > td:nth-child(2)')->text());
                    $b = str_replace(',', '', $tab1->filter('#\32 21 > td:nth-child(3)')->text());
                    $c = str_replace(',', '', $tab1->filter('#\32 21 > td:nth-child(4)')->text());
                    $d = str_replace(',', '', $tab1->filter('#\32 21 > td:nth-child(5)')->text());
                    array_push($data[7], $a, $b, $c, $d);
                }
                

                // // 8 => ["Tài sản cố định"],
                if($tab1->filter('#\32 20 > td:nth-child(2)')->count() > 0) {
                    $a = str_replace(',', '', $tab1->filter('#\32 20 > td:nth-child(2)')->text());
                    $b = str_replace(',', '', $tab1->filter('#\32 20 > td:nth-child(3)')->text());
                    $c = str_replace(',', '', $tab1->filter('#\32 20 > td:nth-child(4)')->text());
                    $d = str_replace(',', '', $tab1->filter('#\32 20 > td:nth-child(5)')->text());
                    array_push($data[8], $a, $b, $c, $d);
                }
                

                // // 9 => ["Tổng lợi nhuận kế toán trước thuế"], 
                if($tab2->filter('#\35 0 > td:nth-child(2)')->count() > 0) {
                    $a = str_replace(',', '', $tab2->filter('#\35 0 > td:nth-child(2)')->text());
                    $b = str_replace(',', '', $tab2->filter('#\35 0 > td:nth-child(3)')->text());
                    $c = str_replace(',', '', $tab2->filter('#\35 0 > td:nth-child(4)')->text());
                    $d = str_replace(',', '', $tab2->filter('#\35 0 > td:nth-child(5)')->text());
                    array_push($data[9], $a, $b, $c, $d);
                }
                

                // // 10 => ["Chi phí thuế TNDN hiện hành"],
                if($tab2->filter('#\35 1 > td:nth-child(2)')->count() > 0) {
                    $a = str_replace(',', '', $tab2->filter('#\35 1 > td:nth-child(2)')->text());
                    $b = str_replace(',', '', $tab2->filter('#\35 1 > td:nth-child(3)')->text());
                    $c = str_replace(',', '', $tab2->filter('#\35 1 > td:nth-child(4)')->text());
                    $d = str_replace(',', '', $tab2->filter('#\35 1 > td:nth-child(5)')->text());
                    array_push($data[10], $a, $b, $c, $d);
                }
                

                // // 11 => ["Lợi nhuận sau thuế thu nhập doanh nghiệp"],
                if($tab2->filter('#\36 0 > td:nth-child(2)')->count() > 0) {
                    $a = str_replace(',', '', $tab2->filter('#\36 0 > td:nth-child(2)')->text());
                    $b = str_replace(',', '', $tab2->filter('#\36 0 > td:nth-child(3)')->text());
                    $c = str_replace(',', '', $tab2->filter('#\36 0 > td:nth-child(4)')->text());
                    $d = str_replace(',', '', $tab2->filter('#\36 0 > td:nth-child(5)')->text());
                    array_push($data[11], $a, $b, $c, $d);
                }
                

                // // 12 => ["Lợi nhuận sau thuế của công ty mẹ"],
                if($tab2->filter('#\36 2 > td:nth-child(2)')->count() > 0) {
                    $a = str_replace(',', '', $tab2->filter('#\36 2 > td:nth-child(2)')->text());
                    $b = str_replace(',', '', $tab2->filter('#\36 2 > td:nth-child(3)')->text());
                    $c = str_replace(',', '', $tab2->filter('#\36 2 > td:nth-child(4)')->text());
                    $d = str_replace(',', '', $tab2->filter('#\36 2 > td:nth-child(5)')->text());
                    array_push($data[12], $a, $b, $c, $d);
                }
               

                // // 13 => ["Chi phí lãi vay"]
                if($tab2->filter('#\32 3 > td:nth-child(2)')->count() > 0) {
                    $a = str_replace(',', '', $tab2->filter('#\32 3 > td:nth-child(2)')->text());
                    $b = str_replace(',', '', $tab2->filter('#\32 3 > td:nth-child(3)')->text());
                    $c = str_replace(',', '', $tab2->filter('#\32 3 > td:nth-child(4)')->text());
                    $d = str_replace(',', '', $tab2->filter('#\32 3 > td:nth-child(5)')->text());
                    array_push($data[13], $a, $b, $c, $d);
                }
                
            }
            //calc 2
            foreach($data[0] as $index => $val) {
                if($index > 0) {
                    $a = str_replace(',', '', $data[0][$index]);
                    $b = str_replace(',', '', $data[1][$index]);
                    $data[2][$index] = $b ? ((float)$a / (float)$b) . '' : 0;
                    // $data[2][$index] = str_replace('.', ',', $data[2][$index]);
                }
            }
            // return $data;
            return Excel::download(new ReportExport($code, $head, $data), 'Report_'.$code.'_Horizontal_'.date('YmdHis').'.xlsx');
        }catch(\InvalidArgumentException $e) {
            abort(404);
        }
    }

    public function exportAll(Request $request) {
        set_time_limit(0);
        $years = [];
        for($i = $request->year_from; $i<=$request->year_to; $i++) {
            array_push($years, $i);
        }
        if($request->horizontal) {
            return $this->exportHorizontalAll($request, $years);
        }
        return $this->exportVerticalAll($request, $years);
    }

    public function exportAll2(Request $request) {
        set_time_limit(0);
        $request->year_from = 2010;
        $request->year_to = 2020;
        $request->request->add(['year_from' => 2010, 'year_to' => 2020]);

        $years = [];
        for($i = $request->year_from; $i<=$request->year_to; $i++) {
            array_push($years, $i);
        }
        return $this->exportHorizontalAll3($request, $years);
    }

    public function exportHorizontalAll3(Request $request, $years) {
        // try{
            $client = new Client();
            // $code = $request->code;
            $codes = $this->getCodes();

            foreach($codes as $key => $code){
                $data = [
                    0 => ["ROA"],
                    1 => ["TĂNG TRƯỞNG NỢ"]
                ];
    
                $head = [''];
                $filePath = 'report/all2/horizontal/'.$request->year_from.'-'.$request->year_to.'/Report_'.$code.'_Horizontal'.'.xlsx';
                if(Storage::exists($filePath)) {
                    continue;
                }
                foreach ($years as $year) {

                    $tab1 = $client->request('GET', 'https://s.cafef.vn/bao-cao-tai-chinh/'.$code.'/BSheet/'.$year.'/4/0/0/ket-qua-hoat-dong-kinh-doanh-cong-ty-co-phan-dau-tu-the-gioi-di-dong.chn');
                    $tab2 = $client->request('GET', 'https://s.cafef.vn/bao-cao-tai-chinh/'.$code.'/IncSta/'.$year.'/4/0/0/ket-qua-hoat-dong-kinh-doanh-cong-ty-co-phan-dau-tu-the-gioi-di-dong.chn');

                    //head
                    $a = $tab1->filter('#tblGridData td:nth-child(2)')->text();
                    $b = $tab1->filter('#tblGridData td:nth-child(3)')->text();
                    $c = $tab1->filter('#tblGridData td:nth-child(4)')->text();
                    $d = $tab1->filter('#tblGridData td:nth-child(5)')->text();
                    array_push($head, $a, $b, $c, $d);
    
                    // 0 
                    $aa1 = $tab2->filter('#\36 0 > td:nth-child(2)')->count() > 0 ? str_replace(',', '', $tab2->filter('#\36 0 > td:nth-child(2)')->text()) : 0;
                    $aa2 = $tab1->filter('#\30 01 > td:nth-child(2)')->count() > 0 ? str_replace(',', '', $tab1->filter('#\30 01 > td:nth-child(2)')->text()) : 0;
                    $a = is_numeric($aa1) && is_numeric($aa2) && $aa2!=0 ? $aa1/$aa2 : 0;

                    $bb1 = $tab2->filter('#\36 0 > td:nth-child(3)')->count() > 0 ? str_replace(',', '', $tab2->filter('#\36 0 > td:nth-child(3)')->text()) : 0;
                    $bb2 = $tab1->filter('#\30 01 > td:nth-child(3)')->count() > 0 ? str_replace(',', '', $tab1->filter('#\30 01 > td:nth-child(3)')->text()) : 0;
                    $b = is_numeric($bb1) && is_numeric($bb2) && $bb2!=0 ? $bb1/$bb2 : 0;

                    $cc1 = $tab2->filter('#\36 0 > td:nth-child(4)')->count() > 0 ? str_replace(',', '', $tab2->filter('#\36 0 > td:nth-child(4)')->text()) : 0;
                    $cc2 = $tab1->filter('#\30 01 > td:nth-child(4)')->count() > 0 ? str_replace(',', '', $tab1->filter('#\30 01 > td:nth-child(4)')->text()) : 0;
                    $c = is_numeric($cc1) && is_numeric($cc2) && $aa2!=0 ? $cc1/$aa2 : 0;

                    $dd1 = $tab2->filter('#\36 0 > td:nth-child(5)')->count() > 0 ? str_replace(',', '', $tab2->filter('#\36 0 > td:nth-child(5)')->text()) : 0;
                    $dd2 = $tab1->filter('#\30 01 > td:nth-child(5)')->count() > 0 ? str_replace(',', '', $tab1->filter('#\30 01 > td:nth-child(5)')->text()) : 0;
                    $d = is_numeric($dd1) && is_numeric($dd2) && $dd2!=0 ? $dd1/$dd2 : 0;
                    array_push($data[0], $a, $b, $c, $d);

                    
                    $a1 = $tab1->filter('#\33 30 > td:nth-child(2)')->count() > 0 ? str_replace(',', '', $tab1->filter('#\33 30 > td:nth-child(2)')->text()) : '';
                    $b1 = $tab1->filter('#\33 30 > td:nth-child(3)')->count() > 0 ? str_replace(',', '', $tab1->filter('#\33 30 > td:nth-child(3)')->text()) : '';
                    $c1 = $tab1->filter('#\33 30 > td:nth-child(4)')->count() > 0 ? str_replace(',', '', $tab1->filter('#\33 30 > td:nth-child(4)')->text()) : '';
                    $d1 = $tab1->filter('#\33 30 > td:nth-child(5)')->count() > 0 ? str_replace(',', '', $tab1->filter('#\33 30 > td:nth-child(5)')->text()) : '';
                    array_push($data[1], $a1, $b1, $c1, $d1);
                }

                // return $data;
                //calc 2
                $temp = $data[1];
                foreach($temp as $index => $val) {
                    if($index > 0 && $index < count($temp) - 1) {
                        $a = $data[1][$index];
                        $b = $data[1][$index+1];
                        $temp[$index] = is_numeric($a) && is_numeric($b) && $a != 0 ? ($b-$a)/$a : 0;
                    }
                    if($index == count($temp) - 1) {
                        $temp[$index] = '';
                    }
                }
                $data[1] = $temp;
                Excel::store(new ReportExport($code, $head, $data), $filePath);
            }

            return "OK";
            // $this->addToZip('report/horizontal/'.$request->year_from.'-'.$request->year_to);
            // return Storage::disk('public')->download('report/vertical/'.$request->year_from.'-'.$request->year_to.'/Export_All.zip');
        // }catch(\InvalidArgumentException $e) {
        //     abort(404);
        // }
    }

    public function exportHorizontalAll2(Request $request, $years) {
        // try{
            $client = new Client();
            // $code = $request->code;
            $codes = $this->getCodes();

            foreach($codes as $key => $code){
                $data = [
                    0 => ["Doanh thu thuần về bán hàng và cung cấp dịch vụ"],
                    1 => ["Tăng trưởng doanh thu"],
                    2 => ["Thuế thu nhập doanh nghiệp"],
                    3 => ["Lợi nhuận trươc thuế"],
                    4 => ["Thuế thu nhập doanh nghiệp/Lợi nhuận trước thuế"],
                    5 => ["Khấu hao"],
                    6 => ["Lợi nhuận sau thuế thu nhập doanh nghiệp"],
                    7 => ["Dòng tiền"],
                    8 => ["Tài sản cố định hữu hình"],
                    9 => ["Tổng tài sản"], 
                    10 => ["Tính hữu hình của tài sản"],
                    11 => ["Nợ dài hạn"]
                ];
    
                $head = [''];
                $filePath = 'report/all2/horizontal/'.$request->year_from.'-'.$request->year_to.'/Report_'.$code.'_Horizontal'.'.xlsx';
                echo $code."<br>";
                if(Storage::exists($filePath)) {
                    continue;
                }
                foreach ($years as $year) {

                    $tab1 = $client->request('GET', 'https://s.cafef.vn/bao-cao-tai-chinh/'.$code.'/BSheet/'.$year.'/4/0/0/ket-qua-hoat-dong-kinh-doanh-cong-ty-co-phan-dau-tu-the-gioi-di-dong.chn');
                    $tab2 = $client->request('GET', 'https://s.cafef.vn/bao-cao-tai-chinh/'.$code.'/IncSta/'.$year.'/4/0/0/ket-qua-hoat-dong-kinh-doanh-cong-ty-co-phan-dau-tu-the-gioi-di-dong.chn');
                    $tab3 = $client->request('GET', 'https://s.cafef.vn/bao-cao-tai-chinh/'.$code.'/CashFlow/'.$year.'/4/0/0/ket-qua-hoat-dong-kinh-doanh-cong-ty-co-phan-dau-tu-the-gioi-di-dong.chn');

                    //head
                    $a = $tab1->filter('#tblGridData td:nth-child(2)')->text();
                    $b = $tab1->filter('#tblGridData td:nth-child(3)')->text();
                    $c = $tab1->filter('#tblGridData td:nth-child(4)')->text();
                    $d = $tab1->filter('#tblGridData td:nth-child(5)')->text();
                    array_push($head, $a, $b, $c, $d);
    
                    // 0 
                    $a = $tab2->filter('#\31 0 > td:nth-child(2)')->count() > 0 ? str_replace(',', '', $tab2->filter('#\31 0 > td:nth-child(2)')->text()) : '';
                    $b = $tab2->filter('#\31 0 > td:nth-child(3)')->count() > 0 ? str_replace(',', '', $tab2->filter('#\31 0 > td:nth-child(3)')->text()) : '';
                    $c = $tab2->filter('#\31 0 > td:nth-child(4)')->count() > 0 ? str_replace(',', '', $tab2->filter('#\31 0 > td:nth-child(4)')->text()) : '';
                    $d = $tab2->filter('#\31 0 > td:nth-child(5)')->count() > 0 ? str_replace(',', '', $tab2->filter('#\31 0 > td:nth-child(5)')->text()) : '';
                    array_push($data[0], $a, $b, $c, $d);
                    

                    // 1
                    $a1 = "";
                    $b1 = "";
                    $c1 = "";
                    $d1 = "";
                    array_push($data[1], $a1, $b1, $c1, $d1);

                    // 2
                    $a2 = $tab2->filter('#\35 1 > td:nth-child(2)')->count() > 0 ? str_replace(',', '', $tab2->filter('#\35 1 > td:nth-child(2)')->text()) : '';
                    $b2 = $tab2->filter('#\35 1 > td:nth-child(3)')->count() > 0 ? str_replace(',', '', $tab2->filter('#\35 1 > td:nth-child(3)')->text()) : '';
                    $c2 = $tab2->filter('#\35 1 > td:nth-child(4)')->count() > 0 ? str_replace(',', '', $tab2->filter('#\35 1 > td:nth-child(4)')->text()) : '';
                    $d2 = $tab2->filter('#\35 1 > td:nth-child(5)')->count() > 0 ? str_replace(',', '', $tab2->filter('#\35 1 > td:nth-child(5)')->text()) : '';
                    array_push($data[2], $a2, $b2, $c2, $d2);

                    // 3
                    $a3 = $tab2->filter('#\35 0 > td:nth-child(2)')->count() > 0 ? str_replace(',', '', $tab2->filter('#\35 0 > td:nth-child(2)')->text()) : '';
                    $b3 = $tab2->filter('#\35 0 > td:nth-child(3)')->count() > 0 ? str_replace(',', '', $tab2->filter('#\35 0 > td:nth-child(3)')->text()) : '';
                    $c3 = $tab2->filter('#\35 0 > td:nth-child(4)')->count() > 0 ? str_replace(',', '', $tab2->filter('#\35 0 > td:nth-child(4)')->text()) : '';
                    $d3 = $tab2->filter('#\35 0 > td:nth-child(5)')->count() > 0 ? str_replace(',', '', $tab2->filter('#\35 0 > td:nth-child(5)')->text()) : '';
                    array_push($data[3], $a3, $b3, $c3, $d3);

                    // 4
                    $a4 = is_numeric($a2) && is_numeric($a3) && $a3!=0 ? $a2/$a3 : 0;
                    $b4 = is_numeric($b2) && is_numeric($b3) && $b3!=0 ? $b2/$b3 : 0;
                    $c4 = is_numeric($c2) && is_numeric($c3) && $c3!=0 ? $c2/$c3 : 0;
                    $d4 = is_numeric($d2) && is_numeric($d3) && $d3!=0 ? $d2/$d3 : 0;
                    array_push($data[4], $a4, $b4, $c4, $d4);

                    // 5
                    $a5 = $tab3->filter('#\30 2 > td:nth-child(2)')->count() > 0 ? str_replace(',', '', $tab3->filter('#\30 2 > td:nth-child(2)')->text()) : '';
                    $b5 = $tab3->filter('#\30 2 > td:nth-child(3)')->count() > 0 ? str_replace(',', '', $tab3->filter('#\30 2 > td:nth-child(3)')->text()) : '';
                    $c5 = $tab3->filter('#\30 2 > td:nth-child(4)')->count() > 0 ? str_replace(',', '', $tab3->filter('#\30 2 > td:nth-child(4)')->text()) : '';
                    $d5 = $tab3->filter('#\30 2 > td:nth-child(5)')->count() > 0 ? str_replace(',', '', $tab3->filter('#\30 2 > td:nth-child(5)')->text()) : '';
                    array_push($data[5], $a5, $b5, $c5, $d5);

                    // 6
                    $a6 = $tab2->filter('#\36 0 > td:nth-child(2)')->count() > 0 ? str_replace(',', '', $tab2->filter('#\36 0 > td:nth-child(2)')->text()) : '';
                    $b6 = $tab2->filter('#\36 0 > td:nth-child(3)')->count() > 0 ? str_replace(',', '', $tab2->filter('#\36 0 > td:nth-child(3)')->text()) : '';
                    $c6 = $tab2->filter('#\36 0 > td:nth-child(4)')->count() > 0 ? str_replace(',', '', $tab2->filter('#\36 0 > td:nth-child(4)')->text()) : '';
                    $d6 = $tab2->filter('#\36 0 > td:nth-child(5)')->count() > 0 ? str_replace(',', '', $tab2->filter('#\36 0 > td:nth-child(5)')->text()) : '';
                    array_push($data[6], $a6, $b6, $c6, $d6);

                    // 7
                    $a7 = is_numeric($a6) || is_numeric($a5) ? intval($a6)+intval($a5) : 0;
                    $b7 = is_numeric($b6) || is_numeric($b5) ? intval($b6)+intval($b5) : 0;
                    $c7 = is_numeric($c6) || is_numeric($c5) ? intval($c6)+intval($c5) : 0;
                    $d7 = is_numeric($d6) || is_numeric($d5) ? intval($d6)+intval($d5) : 0;
                    array_push($data[7], $a7, $b7, $c7, $d7);

                    // 8
                    $a8 = $tab1->filter('#\32 21 > td:nth-child(2)')->count() > 0 ? str_replace(',', '', $tab1->filter('#\32 21 > td:nth-child(2)')->text()) : '';
                    $b8 = $tab1->filter('#\32 21 > td:nth-child(3)')->count() > 0 ? str_replace(',', '', $tab1->filter('#\32 21 > td:nth-child(3)')->text()) : '';
                    $c8 = $tab1->filter('#\32 21 > td:nth-child(4)')->count() > 0 ? str_replace(',', '', $tab1->filter('#\32 21 > td:nth-child(4)')->text()) : '';
                    $d8 = $tab1->filter('#\32 21 > td:nth-child(5)')->count() > 0 ? str_replace(',', '', $tab1->filter('#\32 21 > td:nth-child(5)')->text()) : '';
                    array_push($data[8], $a8, $b8, $c8, $d8);

                    // 9
                    $a9 = $tab1->filter('#\30 01 > td:nth-child(2)')->count() > 0 ? str_replace(',', '', $tab1->filter('#\30 01 > td:nth-child(2)')->text()) : '';
                    $b9 = $tab1->filter('#\30 01 > td:nth-child(3)')->count() > 0 ? str_replace(',', '', $tab1->filter('#\30 01 > td:nth-child(3)')->text()) : '';
                    $c9 = $tab1->filter('#\30 01 > td:nth-child(4)')->count() > 0 ? str_replace(',', '', $tab1->filter('#\30 01 > td:nth-child(4)')->text()) : '';
                    $d9 = $tab1->filter('#\30 01 > td:nth-child(5)')->count() > 0 ? str_replace(',', '', $tab1->filter('#\30 01 > td:nth-child(5)')->text()) : '';
                    array_push($data[9], $a9, $b9, $c9, $d9);

                    //10
                    $a10 = is_numeric($a8) && is_numeric($a9) && $a9!=0 ? $a8/$a9 : 0;
                    $b10 = is_numeric($b8) && is_numeric($b9) && $b9!=0 ? $b8/$b9 : 0;
                    $c10 = is_numeric($c8) && is_numeric($c9) && $c9!=0 ? $c8/$c9 : 0;
                    $d10 = is_numeric($d8) && is_numeric($d9) && $d9!=0 ? $d8/$d9 : 0;
                    array_push($data[10], $a10, $b10, $c10, $d10);

                    // 11
                    $a11 = $tab1->filter('#\33 30 > td:nth-child(2)')->count() > 0 ? str_replace(',', '', $tab1->filter('#\33 30 > td:nth-child(2)')->text()) : '';
                    $b11 = $tab1->filter('#\33 30 > td:nth-child(3)')->count() > 0 ? str_replace(',', '', $tab1->filter('#\33 30 > td:nth-child(3)')->text()) : '';
                    $c11 = $tab1->filter('#\33 30 > td:nth-child(4)')->count() > 0 ? str_replace(',', '', $tab1->filter('#\33 30 > td:nth-child(4)')->text()) : '';
                    $d11 = $tab1->filter('#\33 30 > td:nth-child(5)')->count() > 0 ? str_replace(',', '', $tab1->filter('#\33 30 > td:nth-child(5)')->text()) : '';
                    array_push($data[11], $a11, $b11, $c11, $d11);
                }

                //calc 2
                foreach($data[0] as $index => $val) {
                    if($index > 0 && $index < count($data[0]) - 1) {
                        $a = $data[0][$index];
                        $b = $data[0][$index+1];
                        $data[1][$index] = is_numeric($a) && is_numeric($b) && $a != 0 ? ($b-$a)/$a : 0;
                    }
                }
                Excel::store(new ReportExport($code, $head, $data), $filePath);
            }

            return "OK";
            // $this->addToZip('report/horizontal/'.$request->year_from.'-'.$request->year_to);
            // return Storage::disk('public')->download('report/vertical/'.$request->year_from.'-'.$request->year_to.'/Export_All.zip');
        // }catch(\InvalidArgumentException $e) {
        //     abort(404);
        // }
    }

    public function exportVerticalAll2(Request $request, $years) {
        // try{
            $client = new Client();

            $codes = $this->getCodes();
            
            foreach($codes as $key => $code){
                $head = [
                    "",
                    "Doanh thu thuần về bán hàng và cung cấp dịch vụ",
                    "Tăng trưởng doanh thu",
                    "Thuế thu nhập doanh nghiệp",
                    "Lợi nhuận trươc thuế",
                    "Thuế thu nhập doanh nghiệp/Lợi nhuận trước thuế",
                    "Khấu hao",
                    "Lợi nhuận sau thuế thu nhập doanh nghiệp",
                    "Dòng tiền",
                    "Tài sản cố định hữu hình",
                    "Tổng tài sản",
                    "Tính hữu hình của tài sản"
                ];
                $data = [];
                $filePath = 'report/all2/vertical/'.$request->year_from.'-'.$request->year_to.'/Report_'.$code.'_Vertical'.'.xlsx';
                if(Storage::exists($filePath)) {
                    continue;
                }
                foreach ($years as $year) {
                    $tab1 = $client->request('GET', 'https://s.cafef.vn/bao-cao-tai-chinh/'.$code.'/BSheet/'.$year.'/4/0/0/ket-qua-hoat-dong-kinh-doanh-cong-ty-co-phan-dau-tu-the-gioi-di-dong.chn');
                    $tab2 = $client->request('GET', 'https://s.cafef.vn/bao-cao-tai-chinh/'.$code.'/IncSta/'.$year.'/4/0/0/ket-qua-hoat-dong-kinh-doanh-cong-ty-co-phan-dau-tu-the-gioi-di-dong.chn');
                    $tab3 = $client->request('GET', 'https://s.cafef.vn/bao-cao-tai-chinh/'.$code.'/CashFlow/'.$year.'/4/0/0/ket-qua-hoat-dong-kinh-doanh-cong-ty-co-phan-dau-tu-the-gioi-di-dong.chn');

                    //quy 1
                    $a = str_replace(',', '', $tab1->filter('#tblGridData td:nth-child(2)')->text());
                    $a0 = $tab2->filter('#\31 0 > td:nth-child(2)')->count() > 0 ? str_replace(',', '', $tab2->filter('#\31 0 > td:nth-child(2)')->text()) : '';
                    $a1 = "";
                    $a2 = $tab2->filter('#\35 1 > td:nth-child(2)')->count() > 0 ? str_replace(',', '', $tab2->filter('#\35 1 > td:nth-child(2)')->text()) : '';
                    $a3 = $tab2->filter('#\35 0 > td:nth-child(2)')->count() > 0 ? str_replace(',', '', $tab2->filter('#\35 0 > td:nth-child(2)')->text()) : '';
                    $a4 = is_numeric($a2) && is_numeric($a3) && $a3!=0 ? $a2/$a3 : 0;
                    $a5 = $tab3->filter('#\30 2 > td:nth-child(2)')->count() > 0 ? str_replace(',', '', $tab3->filter('#\30 2 > td:nth-child(2)')->text()) : '';
                    $a6 = $tab2->filter('#\36 0 > td:nth-child(2)')->count() > 0 ? str_replace(',', '', $tab2->filter('#\36 0 > td:nth-child(2)')->text()) : '';
                    $a7 = is_numeric($a6) && is_numeric($a5) ? $a6+$a5 : 0;
                    $a8 = $tab1->filter('#\32 21 > td:nth-child(2)')->count() > 0 ? str_replace(',', '', $tab1->filter('#\32 21 > td:nth-child(2)')->text()) : '';
                    $a9 = $tab1->filter('#\30 01 > td:nth-child(2)')->count() > 0 ? str_replace(',', '', $tab1->filter('#\30 01 > td:nth-child(2)')->text()) : '';
                    $a10 = is_numeric($a8) && is_numeric($a9) && $a9!=0 ? $a8/$a9 : 0;

                    array_push($data, [$a, $a0, $a1, $a2, $a3, $a4, $a5, $a6, $a7, $a8, $a9, $a10]);

                    //quy 2
                    $a = str_replace(',', '', $tab1->filter('#tblGridData td:nth-child(3)')->text());
                    $a0 = $tab2->filter('#\31 0 > td:nth-child(3)')->count() > 0 ? str_replace(',', '', $tab2->filter('#\31 0 > td:nth-child(3)')->text()) : '';
                    $a1 = "";
                    $a2 = $tab2->filter('#\35 1 > td:nth-child(3)')->count() > 0 ? str_replace(',', '', $tab2->filter('#\35 1 > td:nth-child(3)')->text()) : '';
                    $a3 = $tab2->filter('#\35 0 > td:nth-child(3)')->count() > 0 ? str_replace(',', '', $tab2->filter('#\35 0 > td:nth-child(3)')->text()) : '';
                    $a4 = is_numeric($a2) && is_numeric($a3) && $a3!=0 ? $a2/$a3 : 0;
                    $a5 = $tab3->filter('#\30 2 > td:nth-child(3)')->count() > 0 ? str_replace(',', '', $tab3->filter('#\30 2 > td:nth-child(3)')->text()) : '';
                    $a6 = $tab2->filter('#\36 0 > td:nth-child(3)')->count() > 0 ? str_replace(',', '', $tab2->filter('#\36 0 > td:nth-child(3)')->text()) : '';
                    $a7 = is_numeric($a6) && is_numeric($a5) ? $a6+$a5 : 0;
                    $a8 = $tab1->filter('#\32 21 > td:nth-child(3)')->count() > 0 ? str_replace(',', '', $tab1->filter('#\32 21 > td:nth-child(3)')->text()) : '';
                    $a9 = $tab1->filter('#\30 01 > td:nth-child(3)')->count() > 0 ? str_replace(',', '', $tab1->filter('#\30 01 > td:nth-child(3)')->text()) : '';
                    $a10 = is_numeric($a8) && is_numeric($a9) && $a9!=0 ? $a8/$a9 : 0;

                    array_push($data, [$a, $a0, $a1, $a2, $a3, $a4, $a5, $a6, $a7, $a8, $a9, $a10]);

                    // //quy 3
                    $a = str_replace(',', '', $tab1->filter('#tblGridData td:nth-child(4)')->text());
                    $a0 = $tab2->filter('#\31 0 > td:nth-child(4)')->count() > 0 ? str_replace(',', '', $tab2->filter('#\31 0 > td:nth-child(4)')->text()) : '';
                    $a1 = "";
                    $a2 = $tab2->filter('#\35 1 > td:nth-child(4)')->count() > 0 ? str_replace(',', '', $tab2->filter('#\35 1 > td:nth-child(4)')->text()) : '';
                    $a3 = $tab2->filter('#\35 0 > td:nth-child(4)')->count() > 0 ? str_replace(',', '', $tab2->filter('#\35 0 > td:nth-child(4)')->text()) : '';
                    $a4 = is_numeric($a2) && is_numeric($a3) && $a3!=0 ? $a2/$a3 : 0;
                    $a5 = $tab3->filter('#\30 2 > td:nth-child(4)')->count() > 0 ? str_replace(',', '', $tab3->filter('#\30 2 > td:nth-child(4)')->text()) : '';
                    $a6 = $tab2->filter('#\36 0 > td:nth-child(4)')->count() > 0 ? str_replace(',', '', $tab2->filter('#\36 0 > td:nth-child(4)')->text()) : '';
                    $a7 = is_numeric($a6) && is_numeric($a5) ? $a6+$a5 : 0;
                    $a8 = $tab1->filter('#\32 21 > td:nth-child(4)')->count() > 0 ? str_replace(',', '', $tab1->filter('#\32 21 > td:nth-child(4)')->text()) : '';
                    $a9 = $tab1->filter('#\30 01 > td:nth-child(4)')->count() > 0 ? str_replace(',', '', $tab1->filter('#\30 01 > td:nth-child(4)')->text()) : '';
                    $a10 = is_numeric($a8) && is_numeric($a9) && $a9!=0 ? $a8/$a9 : 0;

                    array_push($data, [$a, $a0, $a1, $a2, $a3, $a4, $a5, $a6, $a7, $a8, $a9, $a10]);

                    // //quy 4
                    $a = str_replace(',', '', $tab1->filter('#tblGridData td:nth-child(5)')->text());
                    $a0 = $tab2->filter('#\31 0 > td:nth-child(5)')->count() > 0 ? str_replace(',', '', $tab2->filter('#\31 0 > td:nth-child(5)')->text()) : '';
                    $a1 = "";
                    $a2 = $tab2->filter('#\35 1 > td:nth-child(5)')->count() > 0 ? str_replace(',', '', $tab2->filter('#\35 1 > td:nth-child(5)')->text()) : '';
                    $a3 = $tab2->filter('#\35 0 > td:nth-child(5)')->count() > 0 ? str_replace(',', '', $tab2->filter('#\35 0 > td:nth-child(5)')->text()) : '';
                    $a4 = is_numeric($a2) && is_numeric($a3) && $a3!=0 ? $a2/$a3 : 0;
                    $a5 = $tab3->filter('#\30 2 > td:nth-child(5)')->count() > 0 ? str_replace(',', '', $tab3->filter('#\30 2 > td:nth-child(5)')->text()) : '';
                    $a6 = $tab2->filter('#\36 0 > td:nth-child(5)')->count() > 0 ? str_replace(',', '', $tab2->filter('#\36 0 > td:nth-child(5)')->text()) : '';
                    $a7 = is_numeric($a6) && is_numeric($a5) ? $a6+$a5 : 0;
                    $a8 = $tab1->filter('#\32 21 > td:nth-child(5)')->count() > 0 ? str_replace(',', '', $tab1->filter('#\32 21 > td:nth-child(5)')->text()) : '';
                    $a9 = $tab1->filter('#\30 01 > td:nth-child(5)')->count() > 0 ? str_replace(',', '', $tab1->filter('#\30 01 > td:nth-child(5)')->text()) : '';
                    $a10 = is_numeric($a8) && is_numeric($a9) && $a9!=0 ? $a8/$a9 : 0;

                    array_push($data, [$a, $a0, $a1, $a2, $a3, $a4, $a5, $a6, $a7, $a8, $a9, $a10]);
                }
                
                foreach($data as $index => $val) {
                    if($index < count($data) - 1) {
                        $a = $data[$index][1];
                        $b = $data[$index+1][1];
                        $data[$index][2] = is_numeric($a) && is_numeric($b) && $a != 0 ? ($b-$a)/$a : 0;
                    }
                }
                Excel::store(new ReportExport($code, $head, $data), $filePath);
            }

            // add to archive
            return "Add to archive";
            // $this->addToZip('report/vertical/'.$request->year_from.'-'.$request->year_to);
            // return Storage::disk('public')->download('report/vertical/'.$request->year_from.'-'.$request->year_to.'/Export_All.zip');
        // }catch(\InvalidArgumentException $e) {
        //     abort(404);
        // }
    }

    public function exportVerticalAll(Request $request, $years) {
        // try{
            $client = new Client();

            $codes = $this->getCodes();
            
            foreach($codes as $key => $code){
                $head = [
                    "",
                    "Lợi nhuận sau thuế",
                    "Tổng cộng tài sản",
                    "Lợi nhuận sau thuế/tổng tài sản",
                    "Nợ phải trả",
                    "Nợ dài hạn",
                    "Vốn chủ sở hữu",
                    "Hàng tồn kho",
                    "Tài sản cố định hữu hình",
                    "Tài sản cố định",
                    "Tổng lợi nhuận kế toán trước thuế",
                    "Chi phí thuế TNDN hiện hành",
                    "Lợi nhuận sau thuế thu nhập doanh nghiệp",
                    "Lợi nhuận sau thuế của công ty mẹ",
                    "Chi phí lãi vay"
                ];
                $data = [];
                $filePath = 'report/vertical/'.$request->year_from.'-'.$request->year_to.'/Report_'.$code.'_Vertical'.'.xlsx';
                foreach ($years as $year) {
                    if(Storage::exists($filePath)) {
                        break;
                    }
                    $tab1 = $client->request('GET', 'https://s.cafef.vn/bao-cao-tai-chinh/'.$code.'/BSheet/'.$year.'/4/0/0/ket-qua-hoat-dong-kinh-doanh-cong-ty-co-phan-dau-tu-the-gioi-di-dong.chn');
                    $tab2 = $client->request('GET', 'https://s.cafef.vn/bao-cao-tai-chinh/'.$code.'/IncSta/'.$year.'/4/0/0/ket-qua-hoat-dong-kinh-doanh-cong-ty-co-phan-dau-tu-the-gioi-di-dong.chn');

                    //quy 1
                    $a = str_replace(',', '', $tab1->filter('#tblGridData td:nth-child(2)')->text());
                    $a0 = $tab1->filter('#\34 20 > td:nth-child(2)')->count() > 0 ? str_replace(',', '', $tab1->filter('#\34 20 > td:nth-child(2)')->text()) : '';
                    $a1 = $tab1->filter('#\30 01 > td:nth-child(2)')->count() > 0 ? str_replace(',', '', $tab1->filter('#\30 01 > td:nth-child(2)')->text()) : '';
                    $a2 = is_numeric($a0) && is_numeric($a1) && $a1!=0 ? $a0/$a1 : 0;
                    $a3 = $tab1->filter('#\33 00 > td:nth-child(2)')->count() > 0 ? str_replace(',', '', $tab1->filter('#\33 00 > td:nth-child(2)')->text()) : '';
                    $a4 = $tab1->filter('#\33 30 > td:nth-child(2)')->count() > 0 ? str_replace(',', '', $tab1->filter('#\33 30 > td:nth-child(2)')->text()) : '';
                    $a5 = $tab1->filter('#\34 10 > td:nth-child(2)')->count() > 0 ? str_replace(',', '', $tab1->filter('#\34 10 > td:nth-child(2)')->text()) : '';
                    $a6 = $tab1->filter('#\31 40 > td:nth-child(2)')->count() > 0 ? str_replace(',', '', $tab1->filter('#\31 40 > td:nth-child(2)')->text()) : '';
                    $a7 = $tab1->filter('#\32 21 > td:nth-child(2)')->count() > 0 ? str_replace(',', '', $tab1->filter('#\32 21 > td:nth-child(2)')->text()) : '';
                    $a8 = $tab1->filter('#\32 20 > td:nth-child(2)')->count() > 0 ? str_replace(',', '', $tab1->filter('#\32 20 > td:nth-child(2)')->text()) : '';
                    $a9 = $tab2->filter('#\35 0 > td:nth-child(2)')->count() > 0 ? str_replace(',', '', $tab2->filter('#\35 0 > td:nth-child(2)')->text()) : '';
                    $a10 = $tab2->filter('#\35 1 > td:nth-child(2)')->count() > 0 ? str_replace(',', '', $tab2->filter('#\35 1 > td:nth-child(2)')->text()) : '';
                    $a11 = $tab2->filter('#\36 0 > td:nth-child(2)')->count() > 0 ? str_replace(',', '', $tab2->filter('#\36 0 > td:nth-child(2)')->text()) : '';
                    $a12 = $tab2->filter('#\36 2 > td:nth-child(2)')->count() > 0 ? str_replace(',', '', $tab2->filter('#\36 2 > td:nth-child(2)')->text()) : '';
                    $a13 = $tab2->filter('#\32 3 > td:nth-child(2)')->count() > 0 ? str_replace(',', '', $tab2->filter('#\32 3 > td:nth-child(2)')->text()) : '';
                    array_push($data, [$a, $a0, $a1, $a2, $a3, $a4, $a5, $a6, $a7, $a8, $a9, $a10, $a11, $a12, $a13]);

                    //quy 2
                    $a = str_replace(',', '', $tab1->filter('#tblGridData td:nth-child(3)')->text());
                    $a0 = $tab1->filter('#\34 20 > td:nth-child(3)')->count() > 0 ? str_replace(',', '', $tab1->filter('#\34 20 > td:nth-child(3)')->text()) : '';
                    $a1 = $tab1->filter('#\30 01 > td:nth-child(3)')->count() > 0 ? str_replace(',', '', $tab1->filter('#\30 01 > td:nth-child(3)')->text()) : '';
                    $a2 = is_numeric($a0) && is_numeric($a1) && $a1!=0 ? $a0/$a1 : 0;
                    $a3 = $tab1->filter('#\33 00 > td:nth-child(3)')->count() > 0 ? str_replace(',', '', $tab1->filter('#\33 00 > td:nth-child(3)')->text()) : '';
                    $a4 = $tab1->filter('#\33 30 > td:nth-child(3)')->count() > 0 ? str_replace(',', '', $tab1->filter('#\33 30 > td:nth-child(3)')->text()) : '';
                    $a5 = $tab1->filter('#\34 10 > td:nth-child(3)')->count() > 0 ? str_replace(',', '', $tab1->filter('#\34 10 > td:nth-child(3)')->text()) : '';
                    $a6 = $tab1->filter('#\31 40 > td:nth-child(3)')->count() > 0 ? str_replace(',', '', $tab1->filter('#\31 40 > td:nth-child(3)')->text()) : '';
                    $a7 = $tab1->filter('#\32 21 > td:nth-child(3)')->count() > 0 ? str_replace(',', '', $tab1->filter('#\32 21 > td:nth-child(3)')->text()) : '';
                    $a8 = $tab1->filter('#\32 20 > td:nth-child(3)')->count() > 0 ? str_replace(',', '', $tab1->filter('#\32 20 > td:nth-child(3)')->text()) : '';
                    $a9 = $tab2->filter('#\35 0 > td:nth-child(3)')->count() > 0 ? str_replace(',', '', $tab2->filter('#\35 0 > td:nth-child(3)')->text()) : '';
                    $a10 = $tab2->filter('#\35 1 > td:nth-child(3)')->count() > 0 ? str_replace(',', '', $tab2->filter('#\35 1 > td:nth-child(3)')->text()) : '';
                    $a11 = $tab2->filter('#\36 0 > td:nth-child(3)')->count() > 0 ? str_replace(',', '', $tab2->filter('#\36 0 > td:nth-child(3)')->text()) : '';
                    $a12 = $tab2->filter('#\36 2 > td:nth-child(3)')->count() > 0 ? str_replace(',', '', $tab2->filter('#\36 2 > td:nth-child(3)')->text()) : '';
                    $a13 = $tab2->filter('#\32 3 > td:nth-child(3)')->count() > 0 ? str_replace(',', '', $tab2->filter('#\32 3 > td:nth-child(3)')->text()) : '';
                    array_push($data, [$a, $a0, $a1, $a2, $a3, $a4, $a5, $a6, $a7, $a8, $a9, $a10, $a11, $a12, $a13]);

                    //quy 3
                    $a = str_replace(',', '', $tab1->filter('#tblGridData td:nth-child(4)')->text());
                    $a0 = $tab1->filter('#\34 20 > td:nth-child(4)')->count() > 0 ? str_replace(',', '', $tab1->filter('#\34 20 > td:nth-child(4)')->text()) : '';
                    $a1 = $tab1->filter('#\30 01 > td:nth-child(4)')->count() > 0 ? str_replace(',', '', $tab1->filter('#\30 01 > td:nth-child(4)')->text()) : '';
                    $a2 = is_numeric($a0) && is_numeric($a1) && $a1!=0 ? $a0/$a1 : 0;
                    $a3 = $tab1->filter('#\33 00 > td:nth-child(4)')->count() > 0 ? str_replace(',', '', $tab1->filter('#\33 00 > td:nth-child(4)')->text()) : '';
                    $a4 = $tab1->filter('#\33 30 > td:nth-child(4)')->count() > 0 ? str_replace(',', '', $tab1->filter('#\33 30 > td:nth-child(4)')->text()) : '';
                    $a5 = $tab1->filter('#\34 10 > td:nth-child(4)')->count() > 0 ? str_replace(',', '', $tab1->filter('#\34 10 > td:nth-child(4)')->text()) : '';
                    $a6 = $tab1->filter('#\31 40 > td:nth-child(4)')->count() > 0 ? str_replace(',', '', $tab1->filter('#\31 40 > td:nth-child(4)')->text()) : '';
                    $a7 = $tab1->filter('#\32 21 > td:nth-child(4)')->count() > 0 ? str_replace(',', '', $tab1->filter('#\32 21 > td:nth-child(4)')->text()) : '';
                    $a8 = $tab1->filter('#\32 20 > td:nth-child(4)')->count() > 0 ? str_replace(',', '', $tab1->filter('#\32 20 > td:nth-child(4)')->text()) : '';
                    $a9 = $tab2->filter('#\35 0 > td:nth-child(4)')->count() > 0 ? str_replace(',', '', $tab2->filter('#\35 0 > td:nth-child(4)')->text()) : '';
                    $a10 = $tab2->filter('#\35 1 > td:nth-child(4)')->count() > 0 ? str_replace(',', '', $tab2->filter('#\35 1 > td:nth-child(4)')->text()) : '';
                    $a11 = $tab2->filter('#\36 0 > td:nth-child(4)')->count() > 0 ? str_replace(',', '', $tab2->filter('#\36 0 > td:nth-child(4)')->text()) : '';
                    $a12 = $tab2->filter('#\36 2 > td:nth-child(4)')->count() > 0 ? str_replace(',', '', $tab2->filter('#\36 2 > td:nth-child(4)')->text()) : '';
                    $a13 = $tab2->filter('#\32 3 > td:nth-child(4)')->count() > 0 ? str_replace(',', '', $tab2->filter('#\32 3 > td:nth-child(4)')->text()) : '';
                    array_push($data, [$a, $a0, $a1, $a2, $a3, $a4, $a5, $a6, $a7, $a8, $a9, $a10, $a11, $a12, $a13]);

                    //quy 4
                    $a = str_replace(',', '', $tab1->filter('#tblGridData td:nth-child(5)')->text());
                    $a0 = $tab1->filter('#\34 20 > td:nth-child(5)')->count() > 0 ? str_replace(',', '', $tab1->filter('#\34 20 > td:nth-child(5)')->text()) : '';
                    $a1 = $tab1->filter('#\30 01 > td:nth-child(5)')->count() > 0 ? str_replace(',', '', $tab1->filter('#\30 01 > td:nth-child(5)')->text()) : '';
                    $a2 = is_numeric($a0) && is_numeric($a1) && $a1!=0 ? $a0/$a1 : 0;
                    $a3 = $tab1->filter('#\33 00 > td:nth-child(5)')->count() > 0 ? str_replace(',', '', $tab1->filter('#\33 00 > td:nth-child(5)')->text()) : '';
                    $a4 = $tab1->filter('#\33 30 > td:nth-child(5)')->count() > 0 ? str_replace(',', '', $tab1->filter('#\33 30 > td:nth-child(5)')->text()) : '';
                    $a5 = $tab1->filter('#\34 10 > td:nth-child(5)')->count() > 0 ? str_replace(',', '', $tab1->filter('#\34 10 > td:nth-child(5)')->text()) : '';
                    $a6 = $tab1->filter('#\31 40 > td:nth-child(5)')->count() > 0 ? str_replace(',', '', $tab1->filter('#\31 40 > td:nth-child(5)')->text()) : '';
                    $a7 = $tab1->filter('#\32 21 > td:nth-child(5)')->count() > 0 ? str_replace(',', '', $tab1->filter('#\32 21 > td:nth-child(5)')->text()) : '';
                    $a8 = $tab1->filter('#\32 20 > td:nth-child(5)')->count() > 0 ? str_replace(',', '', $tab1->filter('#\32 20 > td:nth-child(5)')->text()) : '';
                    $a9 = $tab2->filter('#\35 0 > td:nth-child(5)')->count() > 0 ? str_replace(',', '', $tab2->filter('#\35 0 > td:nth-child(5)')->text()) : '';
                    $a10 = $tab2->filter('#\35 1 > td:nth-child(5)')->count() > 0 ? str_replace(',', '', $tab2->filter('#\35 1 > td:nth-child(5)')->text()) : '';
                    $a11 = $tab2->filter('#\36 0 > td:nth-child(5)')->count() > 0 ? str_replace(',', '', $tab2->filter('#\36 0 > td:nth-child(5)')->text()) : '';
                    $a12 = $tab2->filter('#\36 2 > td:nth-child(5)')->count() > 0 ? str_replace(',', '', $tab2->filter('#\36 2 > td:nth-child(5)')->text()) : '';
                    $a13 = $tab2->filter('#\32 3 > td:nth-child(5)')->count() > 0 ? str_replace(',', '', $tab2->filter('#\32 3 > td:nth-child(5)')->text()) : '';
                    array_push($data, [$a, $a0, $a1, $a2, $a3, $a4, $a5, $a6, $a7, $a8, $a9, $a10, $a11, $a12, $a13]);
                }
                Excel::store(new ReportExport($code, $head, $data), $filePath);
            }

            // add to archive
            return "Add to archive";
            // $this->addToZip('report/vertical/'.$request->year_from.'-'.$request->year_to);
            // return Storage::disk('public')->download('report/vertical/'.$request->year_from.'-'.$request->year_to.'/Export_All.zip');
        // }catch(\InvalidArgumentException $e) {
        //     abort(404);
        // }
    }

    public function exportHorizontalAll(Request $request, $years) {
        // try{
            $client = new Client();
            // $code = $request->code;
            $codes = $this->getCodes();

            foreach($codes as $key => $code){
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
                    $filePath = 'report/horizontal/'.$request->year_from.'-'.$request->year_to.'/Report_'.$code.'_Horizontal'.'.xlsx';
                    if(Storage::exists($filePath)) {
                        break;
                    }

                    $tab1 = $client->request('GET', 'https://s.cafef.vn/bao-cao-tai-chinh/'.$code.'/BSheet/'.$year.'/4/0/0/ket-qua-hoat-dong-kinh-doanh-cong-ty-co-phan-dau-tu-the-gioi-di-dong.chn');
                    $tab2 = $client->request('GET', 'https://s.cafef.vn/bao-cao-tai-chinh/'.$code.'/IncSta/'.$year.'/4/0/0/ket-qua-hoat-dong-kinh-doanh-cong-ty-co-phan-dau-tu-the-gioi-di-dong.chn');
    
                    //head
                    $a = $tab1->filter('#tblGridData td:nth-child(2)')->text();
                    $b = $tab1->filter('#tblGridData td:nth-child(3)')->text();
                    $c = $tab1->filter('#tblGridData td:nth-child(4)')->text();
                    $d = $tab1->filter('#tblGridData td:nth-child(5)')->text();
                    array_push($head, $a, $b, $c, $d);
    
                    // 0 => ["Lợi nhuận sau thuế"],
                    $a = str_replace(',', '', $tab1->filter('#\34 20 > td:nth-child(2)')->text());
                    $b = str_replace(',', '', $tab1->filter('#\34 20 > td:nth-child(3)')->text());
                    $c = str_replace(',', '', $tab1->filter('#\34 20 > td:nth-child(4)')->text());
                    $d = str_replace(',', '', $tab1->filter('#\34 20 > td:nth-child(5)')->text());
                    array_push($data[0], $a, $b, $c, $d);
                    
    
                    // 1 => ["Tổng cộng tài sản"],
                    $a = str_replace(',', '', $tab1->filter('#\30 01 > td:nth-child(2)')->text());
                    $b = str_replace(',', '', $tab1->filter('#\30 01 > td:nth-child(3)')->text());
                    $c = str_replace(',', '', $tab1->filter('#\30 01 > td:nth-child(4)')->text());
                    $d = str_replace(',', '', $tab1->filter('#\30 01 > td:nth-child(5)')->text());
                    array_push($data[1], $a, $b, $c, $d);
                    
    
                    //2 => ["Lợi nhuận sau thuế/tổng tài sản"],
    
                    //3 => ["NỢ PHẢI TRẢ"],
                    $a = str_replace(',', '', $tab1->filter('#\33 00 > td:nth-child(2)')->text());
                    $b = str_replace(',', '', $tab1->filter('#\33 00 > td:nth-child(3)')->text());
                    $c = str_replace(',', '', $tab1->filter('#\33 00 > td:nth-child(4)')->text());
                    $d = str_replace(',', '', $tab1->filter('#\33 00 > td:nth-child(5)')->text());
                    array_push($data[3], $a, $b, $c, $d);
                    
    
                    //4 => ["Nợ dài hạn"],
                    if($tab1->filter('#\33 30 > td:nth-child(2)')->count() > 0) {
                        $a = str_replace(',', '', $tab1->filter('#\33 30 > td:nth-child(2)')->text());
                        $b = str_replace(',', '', $tab1->filter('#\33 30 > td:nth-child(3)')->text());
                        $c = str_replace(',', '', $tab1->filter('#\33 30 > td:nth-child(4)')->text());
                        $d = str_replace(',', '', $tab1->filter('#\33 30 > td:nth-child(5)')->text());
                        array_push($data[4], $a, $b, $c, $d);
                    }
    
    
                    //5 => ["Vốn chủ sở hữu"],
                    // if(->count() > 0) {}
                    if($tab1->filter('#\34 00 > td:nth-child(2)')->count() > 0) {
                        $a = str_replace(',', '', $tab1->filter('#\34 00 > td:nth-child(2)')->text());
                        $b = str_replace(',', '', $tab1->filter('#\34 10 > td:nth-child(3)')->text());
                        $c = str_replace(',', '', $tab1->filter('#\34 10 > td:nth-child(4)')->text());
                        $d = str_replace(',', '', $tab1->filter('#\34 10 > td:nth-child(5)')->text());
                        array_push($data[5], $a, $b, $c, $d);
                    }
                    
    
                    // // 6 => ["Hàng tồn kho"],
                    if($tab1->filter('#\31 40 > td:nth-child(2)')->count() > 0) {
                        $a = str_replace(',', '', $tab1->filter('#\31 40 > td:nth-child(2)')->text());
                        $b = str_replace(',', '', $tab1->filter('#\31 40 > td:nth-child(3)')->text());
                        $c = str_replace(',', '', $tab1->filter('#\31 40 > td:nth-child(4)')->text());
                        $d = str_replace(',', '', $tab1->filter('#\31 40 > td:nth-child(5)')->text());
                        array_push($data[6], $a, $b, $c, $d);
                    }
    
                    // // 7 => ["Tài sản cố định hữu hình"],
                    if($tab1->filter('#\32 21 > td:nth-child(2)')->count() > 0) {
                        $a = str_replace(',', '', $tab1->filter('#\32 21 > td:nth-child(2)')->text());
                        $b = str_replace(',', '', $tab1->filter('#\32 21 > td:nth-child(3)')->text());
                        $c = str_replace(',', '', $tab1->filter('#\32 21 > td:nth-child(4)')->text());
                        $d = str_replace(',', '', $tab1->filter('#\32 21 > td:nth-child(5)')->text());
                        array_push($data[7], $a, $b, $c, $d);
                    }
                    
    
                    // // 8 => ["Tài sản cố định"],
                    if($tab1->filter('#\32 20 > td:nth-child(2)')->count() > 0) {
                        $a = str_replace(',', '', $tab1->filter('#\32 20 > td:nth-child(2)')->text());
                        $b = str_replace(',', '', $tab1->filter('#\32 20 > td:nth-child(3)')->text());
                        $c = str_replace(',', '', $tab1->filter('#\32 20 > td:nth-child(4)')->text());
                        $d = str_replace(',', '', $tab1->filter('#\32 20 > td:nth-child(5)')->text());
                        array_push($data[8], $a, $b, $c, $d);
                    }
                    
    
                    // // 9 => ["Tổng lợi nhuận kế toán trước thuế"], 
                    if($tab2->filter('#\35 0 > td:nth-child(2)')->count() > 0) {
                        $a = str_replace(',', '', $tab2->filter('#\35 0 > td:nth-child(2)')->text());
                        $b = str_replace(',', '', $tab2->filter('#\35 0 > td:nth-child(3)')->text());
                        $c = str_replace(',', '', $tab2->filter('#\35 0 > td:nth-child(4)')->text());
                        $d = str_replace(',', '', $tab2->filter('#\35 0 > td:nth-child(5)')->text());
                        array_push($data[9], $a, $b, $c, $d);
                    }
                    
    
                    // // 10 => ["Chi phí thuế TNDN hiện hành"],
                    if($tab2->filter('#\35 1 > td:nth-child(2)')->count() > 0) {
                        $a = str_replace(',', '', $tab2->filter('#\35 1 > td:nth-child(2)')->text());
                        $b = str_replace(',', '', $tab2->filter('#\35 1 > td:nth-child(3)')->text());
                        $c = str_replace(',', '', $tab2->filter('#\35 1 > td:nth-child(4)')->text());
                        $d = str_replace(',', '', $tab2->filter('#\35 1 > td:nth-child(5)')->text());
                        array_push($data[10], $a, $b, $c, $d);
                    }
                    
    
                    // // 11 => ["Lợi nhuận sau thuế thu nhập doanh nghiệp"],
                    if($tab2->filter('#\36 0 > td:nth-child(2)')->count() > 0) {
                        $a = str_replace(',', '', $tab2->filter('#\36 0 > td:nth-child(2)')->text());
                        $b = str_replace(',', '', $tab2->filter('#\36 0 > td:nth-child(3)')->text());
                        $c = str_replace(',', '', $tab2->filter('#\36 0 > td:nth-child(4)')->text());
                        $d = str_replace(',', '', $tab2->filter('#\36 0 > td:nth-child(5)')->text());
                        array_push($data[11], $a, $b, $c, $d);
                    }
                    
    
                    // // 12 => ["Lợi nhuận sau thuế của công ty mẹ"],
                    if($tab2->filter('#\36 2 > td:nth-child(2)')->count() > 0) {
                        $a = str_replace(',', '', $tab2->filter('#\36 2 > td:nth-child(2)')->text());
                        $b = str_replace(',', '', $tab2->filter('#\36 2 > td:nth-child(3)')->text());
                        $c = str_replace(',', '', $tab2->filter('#\36 2 > td:nth-child(4)')->text());
                        $d = str_replace(',', '', $tab2->filter('#\36 2 > td:nth-child(5)')->text());
                        array_push($data[12], $a, $b, $c, $d);
                    }
                   
    
                    // // 13 => ["Chi phí lãi vay"]
                    if($tab2->filter('#\32 3 > td:nth-child(2)')->count() > 0) {
                        $a = str_replace(',', '', $tab2->filter('#\32 3 > td:nth-child(2)')->text());
                        $b = str_replace(',', '', $tab2->filter('#\32 3 > td:nth-child(3)')->text());
                        $c = str_replace(',', '', $tab2->filter('#\32 3 > td:nth-child(4)')->text());
                        $d = str_replace(',', '', $tab2->filter('#\32 3 > td:nth-child(5)')->text());
                        array_push($data[13], $a, $b, $c, $d);
                    }
                    
                }

                //calc 2
                foreach($data[0] as $index => $val) {
                    if($index > 0) {
                        $a = str_replace(',', '', $data[0][$index]);
                        $b = str_replace(',', '', $data[1][$index]);
                        $data[2][$index] = $b ? ((float)$a / (float)$b) . '' : 0;
                        // $data[2][$index] = str_replace('.', ',', $data[2][$index]);
                    }
                }
                Excel::store(new ReportExport($code, $head, $data), $filePath);
                // return $data;
            }

            $this->addToZip('report/horizontal/'.$request->year_from.'-'.$request->year_to);
            return Storage::disk('public')->download('report/vertical/'.$request->year_from.'-'.$request->year_to.'/Export_All.zip');
        // }catch(\InvalidArgumentException $e) {
        //     abort(404);
        // }
    }

    public function addToZip($pathOfFiles) {
        //add to zip file
        $path = $pathOfFiles;
        $zipFile = 'Export_All.zip';
        $zip = new ZipArchive();

        // make folder
        if (!file_exists(Storage::disk('public')->path($path))) {
            mkdir(Storage::disk('public')->path($path), 0777, true);
        }
        $zip->open(Storage::disk('public')->path($path .'/'.$zipFile), ZipArchive::CREATE);

        $fileInFolder = Storage::disk('public')->allFiles($pathOfFiles);
        foreach ($fileInFolder as $key => $file){
            $fileNameDetail = basename($file);
            $zip->addFile(Storage::disk('public')->path($file), $fileNameDetail);
        }
        $zip->close();
    }

    public function getCodes() {
        return [
            "AAA",
            "AAM",
            "ABS",
            "ABT",
            "ACC",
            "ACL",
            "ADS",
            "AGG",
            "AGM",
            "AMD",
            "ANV",
            "APC",
            "APG",
            "APH",
            "ASG",
            "ASM",
            "ASP",
            "AST",
            "ATG",
            "BBC",
            "BCE",
            "BCG",
            "BCM",
            "BFC",
            "BHN",
            "BKG",
            "BMC",
            "BMP",
            "BRC",
            "BTP",
            "BTT",
            "BWE",
            "C32",
            "C47",
            "CAV",
            "CCI",
            "CCL",
            "CDC",
            "CEE",
            "CHP",
            "CIG",
            "CII",
            "CKG",
            "CLC",
            "CLG",
            "CLL",
            "CLW",
            "CMG",
            "CMV",
            "CMX",
            "CNG",
            "COM",
            "CRC",
            "CRE",
            "CSM",
            "CSV",
            "CTD",
            "CTF",
            "CTG",
            "CTI",
            "CTS",
            "CVT",
            "D2D",
            "DAG",
            "DAH",
            "DAT",
            "DBC",
            "DBD",
            "DBT",
            "DC4",
            "DCL",
            "DCM",
            "DGC",
            "DGW",
            "DHA",
            "DHC",
            "DHG",
            "DHM",
            "DIG",
            "DLG",
            "DMC",
            "DPG",
            "DPM",
            "DPR",
            "DQC",
            "DRC",
            "DRH",
            "DRL",
            "DSN",
            "DTA",
            "DTL",
            "DTT",
            "DVP",
            "DXG",
            "DXV",
            "ELC",
            "EMC",
            "EVE",
            "EVG",
            "FCM",
            "FCN",
            "FDC",
            "FIR",
            "FIT",
            "FLC",
            "FMC",
            "FPT",
            "FRT",
            "FTM",
            "GAB",
            "GAS",
            "GDT",
            "GEG",
            "GEX",
            "GIL",
            "GMC",
            "GMD",
            "GSP",
            "GTA",
            "GTN",
            "GVR",
            "HAG",
            "HAH",
            "HAI",
            "HAP",
            "HAR",
            "HAS",
            "HAX",
            "HBC",
            "HCD",
            "HDC",
            "HDG",
            "HHS",
            "HID",
            "HII",
            "HMC",
            "HNG",
            "HOT",
            "HPG",
            "HPX",
            "HQC",
            "HRC",
            "HSG",
            "HSL",
            "HT1",
            "HTI",
            "HTL",
            "HTN",
            "HTV",
            "HU1",
            "HU3",
            "HUB",
            "HVH",
            "HVN",
            "HVX",
            "IBC",
            "ICT",
            "IDI",
            "IJC",
            "ILB",
            "IMP",
            "ITA",
            "ITC",
            "ITD",
            "JVC",
            "KBC",
            "KDC",
            "KDH",
            "KHP",
            "KMR",
            "KOS",
            "KPF",
            "KSB",
            "L10",
            "LAF",
            "LBM",
            "LCG",
            "LCM",
            "LDG",
            "LEC",
            "LGC",
            "LGL",
            "LHG",
            "LIX",
            "LM8",
            "LSS",
            "MCG",
            "MCP",
            "MDG",
            "MHC",
            "MSH",
            "MSN",
            "MWG",
            "NAF",
            "NAV",
            "NBB",
            "NCT",
            "NHH",
            "NKG",
            "NLG",
            "NNC",
            "NSC",
            "NT2",
            "NTL",
            "NVL",
            "NVT",
            "OGC",
            "OPC",
            "PAC",
            "PAN",
            "PC1",
            "PDR",
            "PET",
            "PGC",
            "PGD",
            "PHC",
            "PHR",
            "PIT",
            "PJT",
            "PLP",
            "PLX",
            "PME",
            "PNC",
            "PNJ",
            "POM",
            "POW",
            "PPC",
            "PSH",
            "PTB",
            "PTC",
            "PTL",
            "PVD",
            "PVT",
            "PXI",
            "PXS",
            "PXT",
            "QBS",
            "QCG",
            "RAL",
            "RDP",
            "REE",
            "RIC",
            "ROS",
            "S4A",
            "SAB",
            "SAM",
            "SAV",
            "SBA",
            "SBT",
            "SBV",
            "SC5",
            "SCD",
            "SCR",
            "SCS",
            "SFC",
            "SFG",
            "SFG",
            "SFI",
            "SGN",
            "SGR",
            "SGT",
            "SHA",
            "SHI",
            "SHP",
            "SII",
            "SJD",
            "SJF",
            "SJS",
            "SKG",
            "SMA",
            "SMB",
            "SMC",
            "SPM",
            "SRC",
            "SRF",
            "SSC",
            "ST8",
            "STG",
            "STK",
            "SVC",
            "SVI",
            "SVT",
            "SZC",
            "SZL",
            "TAC",
            "TBC",
            "TCD",
            "TCH",
            "TCL",
            "TCM",
            "TCO",
            "TCR",
            "TCT",
            "TDC",
            "TDG",
            "TDH",
            "TDM",
            "TDP",
            "TDW",
            "TEG",
            "TGG",
            "THG",
            "THI",
            "TIP",
            "TIX",
            "TLD",
            "TLG",
            "TLH",
            "TMP",
            "TMS",
            "TMT",
            "TN1",
            "TNA",
            "TNC",
            "TNH",
            "TNI",
            "TNT",
            "TPC",
            "TRA",
            "TRC",
            "TS4",
            "TSC",
            "TTA",
            "TTB",
            "TTE",
            "TTF",
            "TV2",
            "TVT",
            "TYA",
            "UDC",
            "UIC",
            "VAF",
            "VCF",
            "VCG",
            "VDP",
            "VFG",
            "VGC",
            "VHC",
            "VHM",
            "VIC",
            "VID",
            "VIP",
            "VIS",
            "VJC",
            "VMD",
            "VNE",
            "VNG",
            "VNL",
            "VNM",
            "VNS",
            "VOS",
            "VPD",
            "VPG",
            "VPH",
            "VPI",
            "VPS",
            "VRC",
            "VRE",
            "VSC",
            "VSH",
            "VSI",
            "VTB",
            "VTO",
            "YBM",
            "YEG"
        ];
    }
}
