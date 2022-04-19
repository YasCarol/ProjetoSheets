<?php

namespace App\Http\Controllers;

use PHPMailer\PHPMailer\PHPMailer;

use function GuzzleHttp\json_encode;

class sheetsController extends Controller
{
    private $values;
    private $spreadsheetId;
    private $response;
    private $range;

    public function enviaEmail()
    {
        try {
            $this->conexaoDados();
            if (empty($this->values[1])) {
                dd("SEM DADOS NA PLANILHA");
            }

            //VERIFICA SE TEM CAMPO VAZIO
            krsort($this->values); // ordena os valores em ordem decrescente
            $keys = $this->values[0];
            unset($this->values[0]);
            foreach ($this->values as $k => $value) {
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
                    $retorno['ERROS']['MENSAGENS'] = $mail->ErrorInfo;
                    $this->backup('ERROS!A:D', $email, $senha, $login);
                    $this->clear($k);
                } else {
                    $retorno['ENVIADOS'][$email] = true;
                    $this->backup('ENVIADOS!A:D', $email, $senha, $login);
                    $this->clear($k);
                }
            }
            return $retorno;
        } catch (Exception $e) { 
            return ["status" => false, "mensagem" => $e->getMessage()];
        }
    }
    public function conexaoDados()
    {
        //CONEXAO COM O GOOGLE SHEETS
        $client = new \Google_Client();
        $client->setApplicationName('Google Sheets and PHP');
        $client->setScopes([\Google_Service_Sheets::SPREADSHEETS]);
        $client->setAccessType('offline');
        $client->setAuthConfig(__DIR__ . '/credentials.json');

        //BUSCANDO OS DADOS DAS PLANILHAS
        $this->service = new \Google_Service_Sheets($client);
        $this->spreadsheetId = array('1QH2jiFMqtKhBKiiVxKi23KEqAyyJPH8S_lmI8sYq2Jk'); // ID DAS PLANILHAS
        foreach ($this->spreadsheetId as $va) {
            $this->range = 'ENVIAR!A:S'; // NOME DA PLANILHA E INTERVALO
            $this->response = $this->service->spreadsheets_values->get($va, $this->range);
            $this->values = $this->response->getValues();
        }
    }


    public function backup($rang, $email, $login, $senha)
    {
        $this->conexaoDados();
        $this->range = $rang;
        $map = [
            'EMAIL' => $email,
            'SENHA' => $login,
            'LOGIN' => $senha,
        ];
        $values = [
            [
                json_encode($map)
            ],
        ];
        $body = new \Google_Service_Sheets_ValueRange([
            'majorDimension' => 'COLUMNS',
            'values' => $values
        ]);
        $params = [
            'valueInputOption' => "USER_ENTERED"
        ];
        $result = $this->service->spreadsheets_values->append($this->spreadsheetId, $this->range, $body, $params);
    }
    public function clear($k)
    {
        $v = $k + 1;
        $this->conexaoDados();
        $this->range = 'ENVIAR!A' . $v . ':C' . $v;
        $requestBody = new \Google_Service_Sheets_ClearValuesRequest();
        $response = $this->service->spreadsheets_values->clear($this->spreadsheetId, $this->range, $requestBody);
    }
}
