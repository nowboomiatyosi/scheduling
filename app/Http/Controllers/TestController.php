<?php

namespace App\Http\Controllers;

use App\Jobs\ExportExcelJob;
use Illuminate\Http\Request;

class TestController extends Controller
{
    public function index(Request $request)
    {
        $dateRangeStart = '2024-02-27';
        $dateRangeEnd = '2024-03-11';
        for ($locationId = 20701001; $locationId <= 20701040; $locationId++) {
            $locationIds[] = $locationId;
        }
        // Dispatch the job
        ExportExcelJob::dispatch($locationIds, $dateRangeStart, $dateRangeEnd);
    }
}
