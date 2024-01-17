<?php

namespace App\Exports;

use App\Models\Role;
use App\Models\Sale;
use App\Models\Product;
use Illuminate\Support\Facades\Auth;
use Maatwebsite\Excel\Concerns\FromArray;
use Maatwebsite\Excel\Concerns\ShouldAutoSize;
use Maatwebsite\Excel\Concerns\WithEvents;
use Maatwebsite\Excel\Concerns\WithHeadings;
use Maatwebsite\Excel\Events\AfterSheet;
// use Illuminate\Http\Request;

class SalesExport implements FromArray, WithHeadings, ShouldAutoSize, WithEvents
{
    /**
     * @return \Illuminate\Support\Collection
     */
    function array(): array
    {
        $role = Auth::user()->roles()->first();
        $view_records = Role::findOrFail($role->id)->inRole('record_view');

        $date = request()->date;

        // Check If User Has Permission View  All Records
        $Sales = Sale::with('details', 'client', 'facture', 'warehouse')
            ->where('deleted_at', '=', null)
            ->where('date', '=', $date)
            ->where(function ($query) use ($view_records) {
                if (!$view_records) {
                    return $query->where('user_id', '=', Auth::user()->id);
                }
            })->orderBy('id', 'DESC')->get();

        if ($Sales->isNotEmpty()) {
            $data = [];
            $sumGrandTotal = 0;
            $sumPaid = 0;
            $sumdue = 0;
            $total_cash = 0;
            // $total_g_c = 0;
            $t_tran = 0;
            // $total_g_t = 0;
            // $sum_y_paid = 0;

            foreach ($Sales as $sale) {

                $item['date'] = date_format($sale->created_at,"Y-m-d H:i:s");
                // $item['created_at'] = date_format($sale['created_at'],"Y-m-d H:i:s");
                $item['Ref'] = $sale->Ref;
                // $item['list'] = '';
                foreach($sale['details'] as $thing){
                    $Products = Product::where('id', '=', $thing->product_id)->get();
                    if ($Products->isNotEmpty()) {
                        // $product_name = $Products['name'];
                    }
                    array_push($data, (object)[
                        'date'=>'',
                        'Ref' => '',
                        'agentname' => '',
                        'client' => '',
                        'list'=>$Products[0]->code.' | '.$Products[0]->name,
                        'qty'=>$thing->quantity,
                        'price'=>$thing->price,
                        'total'=>$thing->total,
                        'statut' => '',
                        'GrandTotal' => '',
                        'Paid' => '',
                        'due' => '',
                        'payment_status' => ''
                    ]);
                    // $item['date']='';
                    // $item['Ref'] = '';
                    // $item['agentname'] = '';
                    // $item['client'] = '';
                    // $item['list']= $Products[0]->code.' | '.$Products[0]->name;
                    // $item['qty']= $thing->quantity;
                    // $item['price']=$thing->price;
                    // $item['total']=$thing->total;
                    // $item['statut'] = '';
                    // $item['GrandTotal'] = '';
                    // $item['Paid'] = '';
                    // $item['due'] = '';
                    // $item['payment_status'] = '';
                    // $item['list'] .= $Products[0]->code.' | '.$Products[0]->name.' | '.$thing->quantity.' | '. $thing->total ."\n";
                }
                // $item['user'] = $sale['user']->username;
                $item['agentname'] = $sale['user']->firstname . ' ' . $sale['user']->lastname;
                $item['client'] = $sale['client']->name;
                $item['list']= '';
                $item['qty']= '';
                $item['price']= '';
                $item['total']='';
                // $item['Reglement'] = $sale->notes;
                $item['statut'] = $sale->statut;
                $item['GrandTotal'] = $sale->GrandTotal;
                $item['Paid'] = $sale->paid_amount;
                $item['due'] = $sale->GrandTotal - $sale->paid_amount;
                $item['payment_status'] = $sale->payment_statut;
                //$array = json_decode(json_encode($sale['facture']), true);
                if ($item['payment_status']=='paid'){
                    $a = $sale['facture'];
                    foreach($a as $prop) {
                        $first_prop = $prop;
                        break; // exits the foreach loop
                    } 
                    $pay_type = $first_prop->Reglement;
                    $item['payment_type'] = $pay_type;

                    $sumGrandTotal += $sale->GrandTotal;
                    $sumPaid += $sale->paid_amount;
                    
                    
                    if($item['payment_status']=="paid"){
                        if ($pay_type == "ເງິນສົດ") {
                            $total_cash += $sale->paid_amount;
                            // $total_g_c += $sale->GrandTotal;
                        } else {
                            $t_tran += $sale->paid_amount;
                            // $total_g_t += $sale->GrandTotal;
                        }      
                    }
                }
                    
                $sumdue += $item['due'];
                $data[] = $item;
            }
            // total
            array_push($data, (object)[
                'date'=>'',
                'Ref' => '',
                'agentname' => '',
                'client' => '',
                'list'=>'',
                'qty'=>'',
                'price'=>'',
                'total'=>'',
                'statut' => 'ລວມ : ',
                'GrandTotal' => $sumGrandTotal,
                'Paid' => $sumPaid,
                'due' => '',
                'payment_status' => ''
            ]);
            // ຍັງຄ້າງຢູ່
            array_push($data, (object)[
                'date'=>'',
                'Ref' => '',
                'agentname' => '',
                'client' => '',
                'list'=>'',
                'qty'=>'',
                'price'=>'',
                'total'=>'',
                'statut' => '',
                'GrandTotal' =>'ຍັງຄ້າງຢູ່ : ',
                'Paid' => $sumdue,
                'due' => '',
                'payment_status' => ''
            ]);
            // total tran
            array_push($data, (object)[
                'date'=>'',
                'Ref' => '',
                'agentname' => '',
                'client' => '',
                'list'=>'',
                'qty'=>'',
                'price'=>'',
                'total'=>'',
                'statut' => '',
                'GrandTotal' => 'ເງິນໂອນ : ',
                'Paid' => $t_tran,
                'due' => '',
                'payment_status' => ''
            ]);
            // total cash
            array_push($data, (object)[
                'date'=>'',
                'Ref' => '',
                'list'=>'',
                'agentname' => '',
                'client' => '',
                'list'=>'',
                'qty'=>'',
                'price'=>'',
                'total'=>'',
                'statut' => '',
                'GrandTotal' =>'ເງິນສົດ : ',
                'Paid' => $total_cash,
                'due' => '',
                'payment_status' => ''
            ]);
            
            

        } else {
            $data = [];
        }
        
        return $data;
    }

    public function registerEvents(): array
    {
        // $spreadsheet->getDefaultStyle()->getFont()->setName('Phetsarath OT');
        return [
            AfterSheet::class => function (AfterSheet $event) {
                // $event->sheet->getDefaultStyle()->getFont()->setName('Phetsarath OT');

                // $spreadsheet->getTheme()
                //     ->setThemeFontName('Phetsarath OT')
                //     ->setMinorFontValues('Arial', 'Arial', 'Arial', []);
                // $spreadsheet->getDefaultStyle()->getFont()->setScheme('minor');

                $cellRange = 'A1:K1'; // All headers
                $celBody = 'B2:K20000';

                
                $event->sheet->getDelegate()->getStyle($cellRange)->getFont()->setSize(14);
                $event->sheet->getDelegate()->getStyle($cellRange)->getFont()->setName("Phetsarath OT");

                $event->sheet->getDelegate()->getStyle($celBody)->getFont()->setName("Phetsarath OT");
                $event->sheet->getDelegate()->getStyle($celBody)->getAlignment()->setWrapText(true);


                $styleArray = [
                    'borders' => [
                        'outline' => [
                            'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THICK,
                            'color' => ['argb' => 'FFFF0000'],
                        ],
                    ],

                    'alignment' => [
                        'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_RIGHT,
                    ],
                ];

            },
        ];

    }

    public function headings(): array
    {
        return [
            'ວັນທີ',
            'ເລກທີບິນ',
            'ຜູ້ຂາຍ',
            'ລູກຄ້າ',
            'ລາຍການຂາຍ',
            'ຈຳນວນ',
            'ລາຄາຂາຍ',
            'ເປັນເງິນ',
            'ສະຖານະ',
            'ລວມທັງຫມົດ',
            'ຈ່າຍແລ້ວ',
            'ຍັງຄ້າງຢູ່',
            'ສະຖານະການຈ່າຍ',
            'ປະເພດຈ່າຍ & ຫມາຍເຫດ'
        ];
    }
}