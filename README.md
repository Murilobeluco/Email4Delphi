# Email4Delphi
![Imagem](https://i.imgur.com/0Shl1zW.png)

# Biblioteca de Envio de Email em Delphi

Esta é uma biblioteca para envio de emails em Delphi. A biblioteca suporta duas engines de envio de email:

- **ICS (Internet Component Suite)** - [ICS](https://wiki.overbyte.eu/wiki/index.php/ICS_Download)
- **INDY** - [INDY](https://github.com/IndySockets/Indy.git)

Por padrão, a biblioteca utiliza a engine **INDY**. Você pode alterar a engine de envio de email usando variáveis de compilação.

## Configuração da Engine

- Para usar a engine **ICS**, defina a variável de compilação `EMAIL4D_ICS`.
- Para usar a engine **INDY**, defina a variável de compilação `EMAIL4D_INDY`.

Se nenhuma variável for definida, a biblioteca ativará a engine **INDY**.


## Futuras Implementações

No futuro, está planejada a implementação de uma nova engine chamada **Synapse**[Synapse](https://github.com/geby/synapse.git) para o envio de email no Lazarus.


## Exemplo de Uso

Para usar a biblioteca em seu projeto Delphi, siga os passos abaixo:

1. Adicione o diretório `Email4Delphi\scr` ao Search Path do seu projeto.
2. Inclua `Email4Delphi` na cláusula `uses` do seu projeto.
3. Aqui está um exemplo de como enviar um email:

```delphi
var
  sendEmail: ISendemail;
begin
sendEmail := TSendemail.New
					  .AddTo('destinatario@dominio.com', 'Nome Destinatário')
					  .From('remetente@dominio.com', 'Nome Remetente')
					  .Subject('Assunto do Email')
					  .AddMessage('Conteúdo do email')
					  .Host('smtp.dominio.com')
					  .TLS(True)
					  .Port(587)
					  .Auth(True)
					  .AddAttachment('C:\Caminho\Para\Arquivo.txt')
					  .UserName('usuario@dominio.com')
					  .Password('senha-secreta')
					  .OnLog(
  procedure(ALog: string)
  begin
	writeln(Format('%s', [ALog]));
  end)
  .Send;
end;
```

## Referências

Esta biblioteca utiliza código da [sendemail](https://github.com/dliocode/sendemail.git).

## Contribuição

Se você deseja contribuir para o projeto, por favor, abra uma issue ou um pull request.

