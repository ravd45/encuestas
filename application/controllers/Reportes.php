<?php
defined('BASEPATH') or exit('No direct script access allowed');

class Reportes extends CI_Controller
{

    public function __construct()
    {
        parent::__construct();
        $this->load->library('Utilerias');
        $this->load->library('My_PHPExcel');
        $this->load->model('Encuesta_model');
        $this->cct = array();
        $this->sesion = Utilerias::get_cct_sesion($this)[0];
    } // __construct()

    function get_reporte_excel() {
    if (Utilerias::verifica_sesion_redirige($this)) {
                    $this->styleArray_encabezado = array(
                        'borders' => array(
                                'allborders' => array(
                                        'style' => PHPExcel_Style_Border::BORDER_THIN
                                )
                        ),
                        'fill' => array(
                            'type'  => PHPExcel_Style_Fill::FILL_SOLID,
                            'color' => array(
                                        'rgb' => 'F4FCFC')
                    ),
                    'font' => array(
                            'name'  => 'Arial',
                            'bold'  => true,
                            'color' => array(
                                    'rgb' => '000000'
                            )
                    ),
                    'alignment' =>  array(
                            'horizontal'=> PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
                            'vertical'  => PHPExcel_Style_Alignment::VERTICAL_CENTER
                    )
                );
                $this->styleArray_contenido = array(
                    'borders' => array(
                            'allborders' => array(
                                    'style' => PHPExcel_Style_Border::BORDER_THIN
                            )
                    ),
                    'fill' => array(
                        'type'  => PHPExcel_Style_Fill::FILL_SOLID
                ),
                'font' => array(
                        'name'  => 'Arial',
                        'color' => array(
                                'rgb' => '000000'
                        )
                )
            );

            $this->styleArray_titulo = array(
                'borders' => array(
                        'allborders' => array(
                                'style' => PHPExcel_Style_Border::BORDER_THIN
                        )
                ),
                'fill' => array(
                    'type'  => PHPExcel_Style_Fill::FILL_SOLID,
                    'color' => array(
                                'rgb' => 'DFF5F5')
            ),
            'font' => array(
                    'name'  => 'Arial',
                    'bold'  => true,
                    'color' => array(
                            'rgb' => '000000'
                    )
            ),
            'alignment' =>  array(
                    'horizontal'=> PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
                    'vertical'  => PHPExcel_Style_Alignment::VERTICAL_CENTER
            )
            );

            $obj_phpexcel = new My_PHPExcel();
            $obj_phpexcel->getActiveSheet()->SetCellValue('A1', 'REPORTE DE CAPTURA DE ENCUESTAS');
            $obj_phpexcel->getActiveSheet()->SetCellValue('A2', 'CCT');
            $obj_phpexcel->getActiveSheet()->SetCellValue('B2', 'Turno');
            $obj_phpexcel->getActiveSheet()->SetCellValue('C2', 'Nombre cct');
            $obj_phpexcel->getActiveSheet()->SetCellValue('D2', 'Nombre');
            $obj_phpexcel->getActiveSheet()->SetCellValue('E2', 'Apellido1');
            $obj_phpexcel->getActiveSheet()->SetCellValue('F2', 'Apellido2');
            $obj_phpexcel->getActiveSheet()->SetCellValue('G2', 'Edad');

            $obj_phpexcel->getActiveSheet()->SetCellValue('H2', 'Domicilio');
            $obj_phpexcel->getActiveSheet()->SetCellValue('I2', 'Municipio');
            $obj_phpexcel->getActiveSheet()->SetCellValue('J2', 'Colonia');
            $obj_phpexcel->getActiveSheet()->SetCellValue('K2', 'Localidad');
            $obj_phpexcel->getActiveSheet()->SetCellValue('L2', 'Telefono');
            $obj_phpexcel->getActiveSheet()->SetCellValue('M2', 'Rezago');
            $arr_datos = $this->Encuesta_model->get_reporte_excel();
            $indice = 3;
            foreach ($arr_datos as $item) {
                    $obj_phpexcel->getActiveSheet()->SetCellValue('A'.$indice, utf8_encode($item['cct']));
                    $obj_phpexcel->getActiveSheet()->SetCellValue('B'.$indice, "  ".utf8_encode($item['turno']));
                    $obj_phpexcel->getActiveSheet()->SetCellValue('C'.$indice, "  ".utf8_encode($item['nombre_ct']));
                    $obj_phpexcel->getActiveSheet()->SetCellValue('D'.$indice, "  ".utf8_encode($item['NOMBRE(S)']));
                    $obj_phpexcel->getActiveSheet()->SetCellValue('E'.$indice, "  ".utf8_encode($item['PRIMER APELLIDO']));
                    $obj_phpexcel->getActiveSheet()->SetCellValue('F'.$indice, "  ".utf8_encode($item['SEGUNDO APELLIDO']));
                    $obj_phpexcel->getActiveSheet()->SetCellValue('G'.$indice, "  ".utf8_encode($item['EDAD(más de 15 años)']));

                    $obj_phpexcel->getActiveSheet()->SetCellValue('H'.$indice, "  ".utf8_encode($item['DOMICILIO(calle y número)']));
                    $obj_phpexcel->getActiveSheet()->SetCellValue('I'.$indice, "  ".utf8_encode($item['nom_municipio']));
                    $obj_phpexcel->getActiveSheet()->SetCellValue('J'.$indice, "  ".utf8_encode($item['COLONIA']));
                    $obj_phpexcel->getActiveSheet()->SetCellValue('K'.$indice, "  ".utf8_encode($item['LOCALIDAD']));
                    $obj_phpexcel->getActiveSheet()->SetCellValue('L'.$indice, "  ".utf8_encode($item['TELÉFONO']));
                    $obj_phpexcel->getActiveSheet()->SetCellValue('M'.$indice, "  ".utf8_encode($item['REZAGO']));
                    $indice++;
            }


            $obj_phpexcel->getActiveSheet()->getColumnDimension('A')->setAutoSize(true);
            $obj_phpexcel->getActiveSheet()->getColumnDimension('B')->setAutoSize(true);
            $obj_phpexcel->getActiveSheet()->getColumnDimension('C')->setAutoSize(true);
            $obj_phpexcel->getActiveSheet()->getColumnDimension('D')->setAutoSize(true);
            $obj_phpexcel->getActiveSheet()->getColumnDimension('E')->setAutoSize(true);
            $obj_phpexcel->getActiveSheet()->getColumnDimension('F')->setAutoSize(true);
            $obj_phpexcel->getActiveSheet()->getColumnDimension('G')->setAutoSize(true);

            $obj_phpexcel->getActiveSheet()->getColumnDimension('H')->setAutoSize(true);
            $obj_phpexcel->getActiveSheet()->getColumnDimension('I')->setAutoSize(true);
            $obj_phpexcel->getActiveSheet()->getColumnDimension('J')->setAutoSize(true);
            $obj_phpexcel->getActiveSheet()->getColumnDimension('K')->setAutoSize(true);
            $obj_phpexcel->getActiveSheet()->getColumnDimension('L')->setAutoSize(true);
            $obj_phpexcel->getActiveSheet()->getColumnDimension('M')->setAutoSize(true);


            $nombre_excel = "reporte_encuestas.xlsx";

            $obj_phpexcel->getActiveSheet()->mergeCells('A1:M1');
            $obj_phpexcel->getActiveSheet()->getStyle('A1:M1')->applyFromArray($this->styleArray_titulo);
            $obj_phpexcel->getActiveSheet()->getStyle('A2:M2')->applyFromArray($this->styleArray_encabezado);
            $obj_phpexcel->getActiveSheet()->getStyle('A3:M'.($indice-1))->applyFromArray($this->styleArray_contenido);
            header('Content-Type: application/vnd.ms-excel');
            header('Content-Disposition: attachment;filename='.$nombre_excel);
            header('Cache-Control: max-age=0');
            $obj_writer=PHPExcel_IOFactory::createWriter($obj_phpexcel,'Excel2007');
            $obj_writer->save('php://output');
            exit;

    }// verifica_sesion_redirige
}// get_reporte_excel()




} // class