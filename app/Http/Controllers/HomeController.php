<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;
use App\Product;
use App\ExcelUtilty\ExportExcel;

class HomeController extends Controller
{
    public function index(){

        $products = Product::get();
        return view('index', compact('products'));
    }

    public function export(Request $request){
        $products = Product::select('name', 'description', 'quantity')->get();
        // return $products;
        $header = ['Tên', 'Mô tả', 'Số lượng'];
        return (new ExportExcel($header, $products->toArray()))->download('Sanpham.xlsx');
    }
}
