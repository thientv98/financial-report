<?php

use App\Http\Controllers\ExportController;
use Illuminate\Support\Facades\Route;

/*
|--------------------------------------------------------------------------
| Web Routes
|--------------------------------------------------------------------------
|
| Here is where you can register web routes for your application. These
| routes are loaded by the RouteServiceProvider within a group which
| contains the "web" middleware group. Now create something great!
|
*/
Route::get('/', [ExportController::class, 'index']);

Route::get('/export', [ExportController::class, 'export']);

Route::get('/all', [ExportController::class, 'all']);

Route::get('/export-all', [ExportController::class, 'exportAll']);

Route::get('/all2', [ExportController::class, 'exportAll2']);

Route::get('/all4', [ExportController::class, 'exportAll4']);