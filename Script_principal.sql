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
