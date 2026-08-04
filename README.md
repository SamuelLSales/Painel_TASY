# Atualização do Painel TASY

## Fluxo

1. O GitHub Actions executa a cada duas horas ou manualmente.
2. `scripts/atualizar_dados.py` baixa e valida a planilha do OneDrive.
3. As quatro abas obrigatórias são convertidas para `data.json`.
4. O workflow cria um commit somente depois da validação.
5. O Netlify publica o novo commit.
6. O painel carrega somente `/data.json` e mostra a data da atualização no rodapé.

## Executar manualmente

No GitHub:

1. Abra **Ações**.
2. Selecione **Sincronizar OneDrive ao vivo**.
3. Clique em **Executar fluxo de trabalho**.
4. Confirme que a etapa de extração informa quantidades maiores que zero.
5. Confirme que foi criado um commit `auto: atualização da planilha do OneDrive`.

Depois, no Netlify, confirme que esse mesmo commit aparece como **Published**.

## Trocar o endereço da planilha

O projeto mantém o endereço atual como valor padrão. Para trocar sem editar o
código, crie um secret no repositório:

```text
ONEDRIVE_URL
```

Caminho:

```text
Configurações → Secrets and variables → Actions → New repository secret
```

O valor precisa ser o link de compartilhamento da planilha. O script adiciona
`download=1`, segue os redirecionamentos e rejeita respostas HTML ou arquivos
que não sejam um XLSX válido.

## Validações de segurança

A atualização é cancelada quando:

- o download não retorna um XLSX;
- alguma aba obrigatória não existe;
- alguma coluna obrigatória foi removida ou renomeada;
- Pessoas, Módulos, Unidades Administrativas ou Unidades Assistenciais ficam vazios.

Nessas situações, o `data.json` anterior não é substituído e o GitHub Actions
fica vermelho, permitindo identificar a falha.

## Abas esperadas

```text
02_Cadastro_Pessoas
03_Cadastro_UA
04_Cadastro_UAssist
05_Cadastro_Modulos
```

## Atenção à publicação

O arquivo `data.json` fica acessível publicamente no endereço do painel. Como
ele contém nomes e MASP, a publicação deve estar autorizada pela FHEMIG ou o
painel deve possuir controle de acesso.
