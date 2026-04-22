USE World_Cup_26;

-- Ajuste na Tabela de Ataque
ALTER TABLE estatisticas_ataque 
MODIFY COLUMN Gols_por_partida DECIMAL(10,2),
MODIFY COLUMN Grandes_chances_de_gol_por_jogo DECIMAL(10,2),
MODIFY COLUMN Grandes_chances_perdidas_por_jogo DECIMAL(10,2),
MODIFY COLUMN Total_de_finalizações_por_jogo DECIMAL(10,2),
MODIFY COLUMN Chutes_certos_por_jogo DECIMAL(10,2),
MODIFY COLUMN Chutes_errados_por_jogo DECIMAL(10,2),
MODIFY COLUMN Chutes_bloqueados_por_jogo DECIMAL(10,2),
MODIFY COLUMN Dribles_certos_por_jogo DECIMAL(10,2),
MODIFY COLUMN Escanteios_por_jogo DECIMAL(10,2),
MODIFY COLUMN Faltas_Tiros_Diretos_por_jogo DECIMAL(10,2);

-- Ajuste na Tabela de Passes
ALTER TABLE estatisticas_passes 
MODIFY COLUMN Posse_de_bola DECIMAL(10,2);

-- Ajuste na Tabela de Defesa
ALTER TABLE estatisticas_defesa 
MODIFY COLUMN Gols_sofridos_por_jogo DECIMAL(10,2),
MODIFY COLUMN Desarmes_por_jogo DECIMAL(10,2),
MODIFY COLUMN Interceptações_por_jogo DECIMAL(10,2),
MODIFY COLUMN Cortes_por_jogo DECIMAL(10,2),
MODIFY COLUMN Defesas_por_jogo DECIMAL(10,2),
MODIFY COLUMN Bolas_recuperadas_por_jogo DECIMAL(10,2);

-- Ajuste na Tabela Outros
ALTER TABLE estatisticas_outros 
MODIFY COLUMN Desarmes_por_partida DECIMAL(10,2),
MODIFY COLUMN Perda_da_posse_de_bola_por_jogo DECIMAL(10,2),
MODIFY COLUMN Laterais_por_jogo DECIMAL(10,2),
MODIFY COLUMN Tiros_de_meta_por_jogo DECIMAL(10,2),
MODIFY COLUMN Impedimentos_por_jogo DECIMAL(10,2),
MODIFY COLUMN Faltas_por_jogo DECIMAL(10,2),
MODIFY COLUMN Cartões_amarelos_por_partida DECIMAL(10,2);