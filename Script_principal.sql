CREATE DATABASE World_Cup_26;

USE 	World_Cup_26;

CREATE TABLE  Seleções (
	id_selecao INT PRIMARY KEY,
	nome VARCHAR(50) NOT NULL,
	sigla CHAR(3),
	confederacao ENUM( 'UEFA', 'CONMEBOL', 'CONCACAF', 'CAF', 'AFC', 'OFC'),
	ranking_fifa INT,
	grupo VARCHAR(1)
);
CREATE TABLE estadios (
id_estadio INT PRIMARY KEY, 
nome_estadio VARCHAR (50),
cidade VARCHAR(50),
país ENUM( 'Estados Unidos', 'México', 'Canadá'),
altitude_metros INT,
fuso_horario INT
);
CREATE TABLE estatisticas_geral (
id INT PRIMARY KEY AUTO_INCREMENT,
    id_selecao INT,
Partidas INT,
Gols_marcados INT,
Gols_sofridos INT,
Assistências INT
);
CREATE TABLE estatisticas_ataque (
id INT PRIMARY KEY AUTO_INCREMENT,
    id_selecao INT,
Gols_por_partida INT,
Gols_de_pênalti INT,
Gols_de_falta INT,
Gols_de_dentro_da_área INT,
Gols_de_fora_da_área INT,
Gols_com_a_perna_esquerda INT,
Gols_com_a_perna_direita INT,
Gols_de_cabeça INT,
Grandes_chances_de_gol_por_jogo INT,
Grandes_chances_perdidas_por_jogo INT,
Total_de_finalizações_por_jogo INT,
Chutes_certos_por_jogo INT,
Chutes_errados_por_jogo INT,
Chutes_bloqueados_por_jogo INT,
Dribles_certos_por_jogo INT,
Escanteios_por_jogo INT,
Faltas_Tiros_Diretos_por_jogo INT,
Finalizações_na_trave INT
);
CREATE TABLE estatisticas_passes (
id INT PRIMARY KEY AUTO_INCREMENT,
id_selecao INT,
Contra_ataques INT,
Posse_de_bola INT,
Passes_certos INT,
Passes_no_próprio_campo INT,
Passes_certos_no_terço_final INT,
Bolas_longas INT,
Cruzamentos_certos INT
);
CREATE TABLE estatisticas_defesa (
id INT PRIMARY KEY AUTO_INCREMENT,
id_selecao INT,
Jogos_sem_sofrer_gols INT,
Gols_sofridos_por_jogo INT,
Desarmes_por_jogo INT,
Interceptações_por_jogo INT,
Cortes_por_jogo INT,
Defesas_por_jogo INT,
Bolas_recuperadas_por_jogo INT,
Erros_que_levaram_à_finalização INT,
Erros_que_levaram_ao_gol INT,
Pênaltis_cometidos INT,
Gols_de_pênalti_concedidos INT,
Tirar_em_cima_da_linha INT
);
CREATE TABLE estatisticas_outros (
id INT PRIMARY KEY AUTO_INCREMENT,
id_selecao INT,
Último_homem_a_desarmar INT,
Desarmes_por_partida INT,
Duelos_ganhos_pelo_chão INT,
Duelos_aéreos_ganhos INT,
Perda_da_posse_de_bola_por_jogo INT,
Laterais_por_jogo INT,
Tiros_de_meta_por_jogo INT,
Impedimentos_por_jogo INT,
Faltas_por_jogo INT,
Cartões_amarelos_por_partida INT,
Cartões_vermelhos INT
);


