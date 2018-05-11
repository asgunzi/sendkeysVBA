# sendkeysVBA

Exemplo de uso de SendKeys no VBA


Um comando que pode ser muito útil, especialmente para automatizar processos repetitivos, é o “SendKeys”. Este comando simplesmente emula o teclado.

A sintaxe é muito simples, algo como:

```vba
Application.SendKeys "Bom dia!", True
```

Para enviar um “Enter”, utilizar o símbolo “~”:

```vba
Application.SendKeys "~", True
```
 

 

Exemplo:

Esta rotina vai abrir o Bloco de Notas, esperar um segundo, dar um Enter e escrever “Bom dia!”.

```vba
Shell "NotePad.exe", vbMaximizedFocus

Application.Wait DateTime.Now + DateTime.TimeValue(“00:00:01”)

Application.SendKeys “~”, True ‘Enter
Application.SendKeys “Bom dia!”, True
```


Resultado esperado:

![](https://ferramentasexcelvba.files.wordpress.com/2018/05/resultadosend01.jpg)

 

Exemplo 2:

Um trote comum no ambiente de trabalho é apertar as teclas CTRL+ALT+Seta para baixo. isto irá virar a tela do computador de cabeça para baixo.

![](https://ferramentasexcelvba.files.wordpress.com/2018/05/inverso.jpg)

Dá para automatizar isto, via

```vba
Application.SendKeys "^%{DOWN}", True</code>
```

IMPORTANTE: Para voltar ao normal, CTRL+ALT+Seta para cima.

Baixe a planilha aqui, https://github.com/asgunzi/sendkeysVBA, para exemplos das funções descritas.

 

Sincronia

O grande problema do SendKeys é sincronizar janelas. Por exemplo, é possível abrir o SAP via macro, e ir navegando e preenchendo campos com SendKeys. Mas o tempo de abrir cada janela pode variar muito, e se um campo falhar por falta de sincronia, tudo vai falhar.

Softwares mais robustos de RPA (Robot Process Automation) têm comandos que esperam as janelas corretas serem carregadas, o que facilita bastante automações complexas.

Recomendo o AutoIt (https://www.autoitscript.com/site/autoit/), para quem quiser um software free de RPA.

 

Anexo

Lista de palavras-chave para o uso com SendKeys.

Fonte: https://bettersolutions.com/vba/macros/sendkeys.htm


https://ferramentasexcelvba.wordpress.com/2018/05/11/send-keys-no-vba/

