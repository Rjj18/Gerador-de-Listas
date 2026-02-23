# 🏫 Gerador Dinâmico de Listas Escolares (MVP)

Um sistema leve, seguro e performático desenvolvido inteiramente no Microsoft Excel (Web/Desktop) para automatizar a geração de listas de presença e relatórios para a gestão pedagógica. 

Desenhado especificamente para a realidade de escolas públicas, este MVP elimina o trabalho manual de filtrar planilhas do EOL (Escola Online) e oferece uma interface amigável (UI) que simula um aplicativo web, sem custos adicionais de licenciamento ou consumo de APIs.



## 🎯 O Problema
Equipes gestoras e coordenadores pedagógicos perdem horas semanais filtrando bases de dados extensas para gerar listas de turmas específicas para os professores. Sistemas complexos frequentemente esbarram na falta de familiaridade técnica dos usuários finais ou em restrições de infraestrutura de TI nas escolas.

## 💡 A Solução
Uma arquitetura de "Frontend/Backend" encapsulada em um único arquivo Excel, utilizando Matrizes Dinâmicas (Dynamic Arrays) e conceitos de UI/UX para navegação intuitiva.

### 🛠️ Arquitetura Técnica

O projeto foi construído separando a interface da camada de dados:

1. **Frontend (A Interface - *View*):** - Uma planilha limpa (sem linhas de grade), contendo um menu suspenso (Input) desenhado para parecer um campo de formulário web.
   - O usuário seleciona a turma (ex: `1A`) e o sistema gera a lista instantaneamente.
   - A planilha possui **Proteção de Interface**, impedindo que os usuários alterem a estrutura ou apaguem fórmulas acidentalmente. Apenas o campo de seleção é desbloqueado.

2. **Backend (A Base de Dados - *Model*):**
   - Uma planilha oculta com **Proteção de Estrutura de Pasta de Trabalho** (via senha), que atua como banco de dados relacional (Single Source of Truth).
   - Recebe os dados brutos exportados do sistema da secretaria (ex: EOL).

3. **Query Engine (O Controlador - *Controller*):**
   - Utiliza a função `=FILTRO` combinada com Álgebra Booleana para consultar o backend com complexidade $O(1)$.
   - Exemplo da lógica de extração:
     `=FILTRO(Tb_Alunos[Nome do Aluno]; (Tb_Alunos[Turma] = "lista_" & N1) * (Tb_Alunos[Situação Aluno] = "ATIVO"); "Nenhum aluno ativo")`

## 🚀 Como Utilizar (Para Professores/Gestão)

A interface foi desenhada seguindo o princípio KISS (*Keep It Simple, Stupid*):
1. Abra o arquivo no Excel Online (Teams/SharePoint).
2. Na caixa cinza, selecione a turma desejada.
3. Confirme se os nomes carregaram (o sistema filtra apenas alunos com status "ATIVO").
4. Pressione `CTRL + P` para gerar o PDF ou enviar direto para a impressora.

## 🔒 Segurança e Manutenção
- **Proteção de Planilha:** Impede a quebra do layout e das fórmulas de matriz dinâmica.
- **Proteção de Pasta de Trabalho:** Bloqueia a reexibição da aba contendo os dados sensíveis de todos os alunos da escola.
- **Atualização:** A coordenação atualiza a aba oculta periodicamente, mantendo a transparência da data da última atualização no rodapé do Frontend.

## 🛣️ Roadmap (Próximos Passos)
Este MVP prova o conceito e resolve a dor imediata. Versões futuras (V2) preveem a migração para a Microsoft Power Platform:
- [ ] **Data Source:** Migração do backend para o SharePoint Lists.
- [ ] **Interface:** Criação do Frontend via Power Apps com checkboxes para seleção dinâmica de colunas (Nome, RM, Data de Nascimento).
- [ ] **Automação:** Implementação de Power Automate para orquestrar a geração de documentos (Word Template para PDF) de forma assíncrona.

---
*Desenvolvido na intersecção entre a Coordenação Pedagógica e a Engenharia da Computação.*
