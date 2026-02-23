# Gerador Dinâmico de Listas Escolares (MVP)

<img width="731" height="169" alt="image" src="https://github.com/user-attachments/assets/180864c9-e071-4474-b577-23c81f2f8b17" />

Um sistema leve, seguro e performático desenvolvido inteiramente no Microsoft Excel (Web/Desktop) para automatizar a geração de listas de presença e relatórios para a gestão pedagógica. 

Desenhado especificamente para a realidade de escolas públicas, este MVP elimina o trabalho manual de filtrar planilhas do EOL (Escola Online) e oferece uma interface amigável (UI) que simula um aplicativo web, sem custos adicionais de licenciamento ou consumo de APIs.

## O Problema
Todo o mês o processo de gerar listas é realizado manualmente, desde a cópia dos dados do EOL até a própria formatação do arquivo. Além de gerar uma enorme demanda de retrabalho são gerados inúmeros arquivos que são usando apenas uma vez.

## A Solução
Uma arquitetura de "Frontend/Backend" encapsulada em um único arquivo Excel, utilizando Matrizes Dinâmicas (Dynamic Arrays) e conceitos de UI/UX para navegação intuitiva.

### Arquitetura Técnica

O projeto foi construído separando a interface da camada de dados:

1. **Frontend (A Interface - *View*):** - Uma planilha limpa (sem linhas de grade), contendo um menu suspenso (Input) desenhado para parecer um campo de formulário web.
   - O usuário seleciona a turma (ex: `1A`) e o sistema gera a lista instantaneamente.
   - A planilha possui **Proteção de Interface**, impedindo que os usuários alterem a estrutura ou apaguem fórmulas acidentalmente. Apenas o campo de seleção é desbloqueado.

2. **Backend (A Base de Dados - *Model*):**
   - Uma planilha oculta com **Proteção de Estrutura de Pasta de Trabalho** (via senha), que atua como banco de dados relacional (Single Source of Truth).
   - Recebe os dados brutos exportados do sistema da secretaria (ex: EOL).

3. **Query Engine (O Controlador - *Controller*):**
   - Utiliza a função `=FILTRO` combinada com Álgebra Booleana para consultar o backend com complexidade $O(1)$.
   - Exemplo da lógica de extração para a lista de presença:
     `=FILTRO(Tb_Alunos[Nome do Aluno]; (Tb_Alunos[Turma] = "lista_" & [célula com o valor da turma]) * (Tb_Alunos[Situação Aluno] = "ATIVO"); "Nenhum aluno ativo")`
     Essa mesma lógica pode ser replicada para "puxar" dados de outras colunas das tabelas.

## Como Utilizar

A interface foi desenhada seguindo o princípio KISS (*Keep It Simple, Stupid*):
1. Abra o arquivo no Excel Online (Teams/SharePoint).
2. Na caixa cinza, selecione a turma desejada.
3. Confirme se os nomes carregaram (o sistema filtra apenas alunos com status "ATIVO").
4. Pressione `CTRL + P` para gerar o PDF ou enviar direto para a impressora.

![Gravação de tela de 2026-02-23 00-14-47](https://github.com/user-attachments/assets/1ce237e3-bda4-4a7f-a150-b1c5171478a0)


## Segurança e Manutenção
- **Proteção de Planilha:** Impede a quebra do layout e das fórmulas de matriz dinâmica.
- **Proteção de Pasta de Trabalho:** Bloqueia a reexibição da aba contendo os dados sensíveis de todos os alunos da escola.
- **Atualização:** A coordenação atualiza a aba oculta periodicamente, mantendo a transparência da data da última atualização no rodapé do Frontend.

## Roadmap (Próximos Passos)
Este MVP prova o conceito e resolve a dor imediata. Versões futuras (V2) preveem a migração para a Microsoft Power Platform:
- [ ] **Data Source:** Migração do backend para o SharePoint Lists.
- [ ] **Interface:** Criação do Frontend via Power Apps com checkboxes para seleção dinâmica de colunas.
- [ ] **Automação:** Implementação de Power Automate para orquestrar a geração de documentos (Word Template para PDF) de forma assíncrona.

## Licença
Este projeto está sob a licença [MIT](https://choosealicense.com/licenses/mit/) - sinta-se livre para usar, modificar e distribuir. É uma iniciativa aberta, ideal para ser replicada e adaptada por outras unidades da rede pública de ensino.

## Agradecimentos e Ferramentas
* Idealização, lógica de negócio e testes em ambiente escolar por Rjj18.
* Estruturação arquitetural, *code review* das matrizes dinâmicas e redação desta documentação realizadas com o suporte de IA (Google Gemini), atuando como Arquiteto de Soluções em *Pair Programming*.
