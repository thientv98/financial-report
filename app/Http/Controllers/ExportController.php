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
            $this->addToZip('report/vertical/'.$request->year_from.'-'.$request->year_to);
            return 'Add to archived';
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
            return 'Add to archived';
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
            "PND",
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
            "PVD",
            "POW",
            "PPC",
            "PSH",
            "PVT",
            "PTB",
            "PTC",
            "PTL",
            "PXI",
            "PXS",
            "PXT",
            "SFG",
            "SFC",
            "SCS",
            "SCR",
            "SCD",
            "SC5",
            "SBV",
            "SBT",
            "SBA",
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
