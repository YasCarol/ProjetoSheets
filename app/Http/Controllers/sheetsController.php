<?php

namespace App\Http\Controllers;

use PHPMailer\PHPMailer\PHPMailer;

use function GuzzleHttp\json_encode;

class sheetsController extends Controller
{
    // public function conexaoDados()
    // {
    //     //CONEXAO COM O GOOGLE SHEETS
    //     $client = new \Google_Client();
    //     $client->setApplicationName('Google Sheets and PHP');
    //     $client->setScopes([\Google_Service_Sheets::SPREADSHEETS]);
    //     $client->setAccessType('offline');
    //     $client->setAuthConfig(__DIR__ . '/credentials.json');
    //     //BUSCANDO OS DADOS DAS PLANILHAS
    //     $service = new \Google_Service_Sheets($client);
    //     $spreadsheetId = array('1QH2jiFMqtKhBKiiVxKi23KEqAyyJPH8S_lmI8sYq2Jk');
    // }
    function Sheets()
    {
        //CONEXAO COM O GOOGLE SHEETS
        $client = new \Google_Client();
        $client->setApplicationName('Google Sheets and PHP');
        $client->setScopes([\Google_Service_Sheets::SPREADSHEETS]);
        $client->setAccessType('offline');
        $client->setAuthConfig(__DIR__ . '/credentials.json');

        //BUSCANDO OS DADOS DAS PLANILHAS
        $service = new \Google_Service_Sheets($client);
        $spreadsheetId = array('1QH2jiFMqtKhBKiiVxKi23KEqAyyJPH8S_lmI8sYq2Jk'); // ID DAS PLANILHAS

        foreach ($spreadsheetId as $va) {
            $range = 'ENVIAR!A:S'; // NOME DA PLANILHA E INTERVALO
            $response = $service->spreadsheets_values->get($va, $range);
            $values = $response->getValues();

            //VERIFICA SE TEM DADOS NA PLANILHA
            if (empty($values[1])) {
                dd("SEM DADOS NA PLANILHA");
            }

            //VERIFICA SE TEM CAMPO VAZIO
            krsort($values); // ordena os valores em ordem decrescente
            $keys = $values[0];
            unset($values[0]);
            foreach ($values as $k => $value) {
                foreach ($keys as $key => $va) {
                    if (empty($value[$key])) {
                        $retorno['TRAVADO'][$value[0]] = $va;
                        return $retorno;
                    }
                }

                //ATRIBUI VALORES
                $email = $value[0];
                $login = $value[1];
                $senha = $value[2];

                //ENVIO DE EMAIL
                $mail = new PHPMailer();
                $mail->isSMTP();
                $mail->Host = 'smtp.gmail.com';
                $mail->SMTPAuth = true;
                $mail->SMTPSecure = 'tls';
                $mail->Username = 'carolineyasmin815@gmail.com';
                $mail->Password = 'yasmin123456';
                $mail->Port = 587;
                $mail->setFrom('yasminazapfy@gmail.com', 'yasmin');
                $mail->addAddress($email);
                $mail->Subject = 'Envio de email';
                $mail->Body = 'Ola, Boa tarde. Segue sua senha email e login ' . $senha . " " . $login . ' para acessar o sistema.';
                if (!$mail->send()) {
                    $retorno['ERRO']['MENSAGEM'] = $mail->ErrorInfo;
                } else {
                    $retorno['ENVIADOS'][$email] = true;
                    $this->backup($email, $senha, $login);
                    $this->clear($k);
                }
            }
        }
        return $retorno;
    }

    public function backup($email, $login, $senha)
    {
        //CONEXAO COM O GOOGLE SHEETS
        $client = new \Google_Client();
        $client->setApplicationName('Google Sheets and PHP');
        $client->setScopes([\Google_Service_Sheets::SPREADSHEETS]);
        $client->setAccessType('offline');
        $client->setAuthConfig(__DIR__ . '/credentials.json');
        $service = new \Google_Service_Sheets($client);
        $range = 'ENVIADOS!A:D';
        $map = [
            'EMAIL' => $email,
            'SENHA' => $login,
            'LOGIN' => $senha,
        ];
        $values = [
            [
                json_encode($map)
            ],
            // Additional rows ...
        ];
        $body = new \Google_Service_Sheets_ValueRange([
            'majorDimension' => 'COLUMNS',
            'values' => $values
        ]);
        $params = [
            'valueInputOption' => "USER_ENTERED"
        ];

        $result = $service->spreadsheets_values->append(
            '1QH2jiFMqtKhBKiiVxKi23KEqAyyJPH8S_lmI8sYq2Jk',
            $range,
            $body,
            $params
        );
    }
    public function clear($k)
    {
        $v = $k + 1;
        //CONEXAO COM O GOOGLE SHEETS
        $client = new \Google_Client();
        $client->setApplicationName('Google Sheets and PHP');
        $client->setScopes([\Google_Service_Sheets::SPREADSHEETS]);
        $client->setAccessType('offline');
        $client->setAuthConfig(__DIR__ . '/credentials.json');
        $service = new \Google_Service_Sheets($client);

        $spreadsheetId = '1QH2jiFMqtKhBKiiVxKi23KEqAyyJPH8S_lmI8sYq2Jk';


        $range = 'ENVIAR!A' . $v . ':C' . $v;

        $requestBody = new \Google_Service_Sheets_ClearValuesRequest();

        $response = $service->spreadsheets_values->clear($spreadsheetId, $range, $requestBody);
    }
}
