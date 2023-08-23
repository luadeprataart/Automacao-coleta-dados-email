# Automacao-coleta-dados-email
Script para automação de tarefas, onde ele realiza interações automatizadas com o serviço de e-mail Outlook por meio do navegador Firefox.

## Configurações do ambiente
* Ativar o ambiente virtual
Navegue até a pasta do projeto e execute o comando:  
 ```shell
 .venv/Scripts/activate
 ```

* Instalando as dependências do projeto
```shell
pip install -r requirements.txt 
```
* Criando o arquivo de config
Crie um arquivo config.py na raiz do projeto e dentro dele cole a seguinte estrutura substituindo os xxx pelos seus dados.
```shell
email = 'xxx'
pass_email = 'xxxx'
Search_email = 'xxxxx'
emergency_contact = '55xxxxxx'
link = 'xxxxxx'
grupo = 'xxxxxxx'
```shell

#### Rodando a aplicação:
```shell
python app_email.py
``` 
