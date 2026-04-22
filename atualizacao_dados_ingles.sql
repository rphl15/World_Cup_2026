USE World_Cup_26;
SET SQL_SAFE_UPDATES = 0;
UPDATE seleções 
SET nome = CASE 
    -- Grupo A
    WHEN nome = 'México' THEN 'Mexico'
    WHEN nome = 'África do Sul' THEN 'South Africa'
    WHEN nome IN ('Córeia do Sul', 'Coreia do Sul') THEN 'South Korea'
    WHEN nome = 'Tchéquia' THEN 'Czech Republic'
    -- Grupo B
    WHEN nome = 'Canadá' THEN 'Canada'
    WHEN nome = 'Bósnia e Herzegovina' THEN 'Bosnia and Herzegovina'
    WHEN nome = 'Catar' THEN 'Qatar'
    WHEN nome = 'Suíça' THEN 'Switzerland'
    -- Grupo C
    WHEN nome = 'Brasil' THEN 'Brazil'
    WHEN nome = 'Marrocos' THEN 'Morocco'
    WHEN nome = 'Escócia' THEN 'Scotland'
    -- Grupo D
    WHEN nome = 'EUA' THEN 'USA'
    WHEN nome = 'Paraguai' THEN 'Paraguay'
    WHEN nome = 'Austrália' THEN 'Australia'
    WHEN nome = 'Turquia' THEN 'Turkey'
    -- Grupo E
    WHEN nome = 'Alemanha' THEN 'Germany'
    WHEN nome = 'Curaçao' THEN 'Curacao'
    WHEN nome = 'Costa do Marfim' THEN 'Ivory Coast'
    WHEN nome = 'Equador' THEN 'Ecuador'
    -- Grupo F
    WHEN nome = 'Holanda' THEN 'Netherlands'
    WHEN nome = 'Japão' THEN 'Japan'
    WHEN nome = 'Suécia' THEN 'Sweden'
    WHEN nome = 'Tunísia' THEN 'Tunisia'
    -- Grupo G
    WHEN nome = 'Bélgica' THEN 'Belgium'
    WHEN nome = 'Egito' THEN 'Egypt'
    WHEN nome = 'Irã' THEN 'Iran'
    WHEN nome = 'Nova Zelândia' THEN 'New Zealand'
    -- Grupo H
    WHEN nome = 'Espanha' THEN 'Spain'
    WHEN nome = 'Cabo Verde' THEN 'Cape Verde'
    WHEN nome = 'Arábia Saudita' THEN 'Saudi Arabia'
    WHEN nome = 'Uruguai' THEN 'Uruguay'
    -- Grupo I
    WHEN nome = 'França' THEN 'France'
    WHEN nome = 'Iraque' THEN 'Iraq'
    WHEN nome = 'Noruega' THEN 'Norway'
    -- Grupo J
    WHEN nome = 'Argélia' THEN 'Algeria'
    WHEN nome = 'Áustria' THEN 'Austria'
    WHEN nome = 'Jordânia' THEN 'Jordan'
    -- Grupo K
    WHEN nome = 'Colômbia' THEN 'Colombia'
    WHEN nome = 'Uzbequistão' THEN 'Uzbekistan'
    WHEN nome = 'RD Congo' THEN 'DR Congo'
    -- Grupo L
    WHEN nome = 'Inglaterra' THEN 'England'
    WHEN nome = 'Croácia' THEN 'Croatia'
    WHEN nome = 'Gana' THEN 'Ghana'
    WHEN nome = 'Panamá' THEN 'Panama'
    ELSE nome 
END;