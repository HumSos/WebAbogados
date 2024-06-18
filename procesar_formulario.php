<?php
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

if ($_SERVER["REQUEST_METHOD"] == "POST") {
    $datos = [
        'primer_apellido' => $_POST['primer_apellido'],
        'segundo_apellido' => $_POST['segundo_apellido'],
        'primer_nombre' => $_POST['primer_nombre'],
        'segundo_nombre' => $_POST['segundo_nombre'],
        'municipio' => $_POST['municipio'],
        'cedula' => $_POST['cedula'],
        'fecha_nacimiento' => $_POST['fecha_nacimiento'],
        'lugar_nacimiento' => $_POST['lugar_nacimiento'],
        'pais_nacimiento' => $_POST['pais_nacimiento'],
        'edad' => $_POST['edad'],
        'estado_civil' => $_POST['estado_civil'],
        'nombre_conyuge' => $_POST['nombre_conyuge'],
        'telefono_conyuge' => $_POST['telefono_conyuge'],
        'telefono_oficina' => $_POST['telefono_oficina'],
        'domicilio_habitacion' => $_POST['domicilio_habitacion'],
        'domicilio_trabajo' => $_POST['domicilio_trabajo'],
        'movil' => $_POST['movil'],
        'cargo' => $_POST['cargo'],
        'email' => $_POST['email'],
        'nro_colegio' => $_POST['nro_colegio'],
        'tomo_colegio' => $_POST['tomo_colegio'],
        'folio_colegio' => $_POST['folio_colegio'],
        'inpreabogado' => $_POST['inpreabogado'],
        'universidad' => $_POST['universidad'],
        'fecha_graduacion' => $_POST['fecha_graduacion'],
        'anios_graduacion' => $_POST['anios_graduacion'],
        'nro_libro' => $_POST['nro_libro'],
        'nro_folio' => $_POST['nro_folio'],
        'estado' => $_POST['estado'],
        'diplomado' => $_POST['diplomado'],
        'post_grado' => $_POST['post_grado'],
        'maestria' => $_POST['maestria'],
        'registro_publico_estado' => $_POST['registro_publico_estado'],
        'bajo_el_nro' => $_POST['bajo_el_nro'],
        'fecha_registro' => $_POST['fecha_registro'],
        'tomo_nro' => $_POST['tomo_nro'],
        'folio_nro' => $_POST['folio_nro'],
        'colegio_abg_inscrito' => $_POST['colegio_abg_inscrito'],
        'fecha_inscripcion' => $_POST['fecha_inscripcion'],
        'nro_tomo_1' => $_POST['nro_tomo_1'],
        'nro_folio_2' => $_POST['nro_folio_2'],
        'nro_colegio2' => $_POST['nro_colegio2'],
        'inpreabogado_nro' => $_POST['inpreabogado_nro'],
        'deportiva' => $_POST['deportiva'],
        'cultural' => $_POST['cultural'],
    ];

    $filePath = 'datos_inscripcion.xlsx';

    if (file_exists($filePath)) {
        $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($filePath);
        $worksheet = $spreadsheet->getActiveSheet();
    } else {
        $spreadsheet = new Spreadsheet();
        $worksheet = $spreadsheet->getActiveSheet();
        $headers = array_keys($datos);
        $row = '1';
        foreach ($headers as $header) {
            $worksheet->setCellValue('A' . $row, $header);
            $row++;
        }
    }

    $row = '1';
    $col = 'B';
    foreach ($datos as $dato) {
        $worksheet->setCellValue('B'. $row, $dato);
        $row++;
    }

    $writer = new Xlsx($spreadsheet);
    $writer->save($filePath);

    echo "Datos guardados correctamente.";
}
?>
