# Atualização do Sistema LIMS - Módulo de Gestão de Equipamentos

## Resumo das Mudanças

O sistema LIMS foi atualizado para integrar um módulo completo de Gestão de Equipamentos e Inovação Tecnológica, baseado nos dados do documento de monitoramento do LINDEF.

## 1. Novas Classes Implementadas

### Classe Base: Equipamento
- Atributos: nome, fabricante, status_operacional, ultima_calibracao, pesquisador_responsavel
- Métodos: registrar_uso, adicionar_dado_medicao, vincular_meta, to_dict

### Classes Específicas por Equipamento

#### RDA (Respirador Digital Aberto) - Vinculado a Emanuel Fernandes Ferreira da Silva Júnior
- registrar_calibracao_seringa(): Registra calibração com seringa
- registrar_problema_engenharia(): Monitora problemas de engenharia

#### USG (Ultrassonografia) - Vinculado a Emanuel Fernandes Ferreira da Silva Júnior
- registrar_analise_imagem(): Análise de imagem (aeração e diafragma)
- registrar_treinamento_confiabilidade(): Treinamentos de confiabilidade

#### EIT (Tomografia de Impedância Elétrica) - Vinculado a Maria Eduarda
- padronizar_uso_estudos_clinicos(): Padronização para estudos clínicos
- registrar_backup_seguro(): Backup seguro de dados

#### EMG/sEMG (Eletromiografia) - Vinculado a Anderson
- registrar_protocolo_higienizacao(): Higienização de equipamentos de monitoração

#### ASL - Vinculado a Maria Eduarda
- adicionar_parametro_patologico(): Tabela de parâmetros de padrões patológicos

#### PowerLab e microFET - Vinculado a Maria Eduarda
- registrar_meta_instrumentacao(): Metas de instrumentação
- registrar_coleta_tcc_ic(): Coletas de TCC/Iniciação Científica

#### Equipamentos Genéricos
- Manovacuômetros, Espirômetros, Ventilômetros, Cateter de Balão Esofágico, Arctic Sun, PEEP Table

## 2. Módulo de Gestão de Equipamentos

### ModuloGestaoEquipamentos
- Inicialização automática de equipamentos
- Controle de acesso baseado em perfis (Docente vs Aluno)
- Captura automática de dados via Excel/CSV
- Geração automática de POPs/MOPs via python-docx
- Validação automática de metas vinculadas a equipamentos

## 3. Integração com Sistema LIMS

### Novos Métodos na Classe SistemaLIMS
- capturar_dados_equipamento_excel/csv(): Captura automática de dados
- gerar_documentacao_equipamento(): Gera POPs/MOPs
- validar_acesso_equipamento(): Controle de acesso
- registrar_uso_equipamento_meta(): Validação de metas
- Métodos específicos para cada equipamento (RDA, USG, EIT, etc.)
- gerar_relatorio_equipamentos_detalhado(): Relatório completo

## 4. Tratamento de Erros e Segurança

- Blocos try/except para falhas de conexão
- Validação de tipos de dados
- Controle de acesso por perfil de usuário
- Logs detalhados de operações

## 5. Vinculação com Pesquisadores (Auditoria 360º)

### Emanuel Fernandes Ferreira da Silva Júnior
- RDA (Calibração com Seringa, Problemas de Engenharia)
- USG (Análise de Imagem, Treinamento de Confiabilidade)
- PEEP Table

### Maria Eduarda
- EIT (Padronização Clínica, Backup Seguro)
- ASL (Tabela de Parâmetros Patológicos)
- PowerLab (Metas de Instrumentação, Coletas TCC/IC)
- microFET (Metas de Instrumentação, Coletas TCC/IC)
- Manovacuômetros, Espirômetros, Ventilômetros

### Anderson
- EMG/sEMG (Protocolos de Higienização)
- Cateter de Balão Esofágico, Arctic Sun

## 6. Menu Interativo Atualizado

Novas opções no menu:
- 17-18: Captura de dados Excel/CSV
- 19: Geração de POP/MOP
- 20: Registro de uso para meta
- 21-30: Operações específicas por equipamento
- 31: Relatório detalhado de equipamentos

## 7. Validação Automática

- Meta de Layane valida automaticamente uso do RDA
- Sistema vincula metas específicas aos equipamentos
- Rastreabilidade completa de uso e calibração

## Dependências Necessárias

- openpyxl: Para captura de dados Excel
- python-docx: Para geração de documentação
- csv: Módulo padrão do Python

## Como Usar

1. Execute o sistema: `python lims_V3.1.2.py`
2. Inicialize o LINDEF (opção 0)
3. Use as opções 17-31 para gestão de equipamentos
4. Gerei relatórios detalhados (opção 31)

O sistema agora garante rastreabilidade completa e automação dos processos analíticos, substituindo o registro manual conforme solicitado.