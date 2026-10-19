-- DONNÉES DE DÉVELOPPEMENT UNIQUEMENT.
-- Les coordonnées sont APPROXIMATIVES / FICTIVES autour du centre d'Abéché et n'ont PAS été
-- relevées sur le terrain (verified = false). À remplacer par la collecte terrain (phase 1 bis).
INSERT INTO zones (name, area) VALUES
  ('Abeche - zone pilote (provisoire)',
   ST_GeogFromText('POLYGON((20.78 13.78, 20.89 13.78, 20.89 13.87, 20.78 13.87, 20.78 13.78))'));

INSERT INTO tariffs (transport_type, base_xaf, per_km_xaf, min_xaf) VALUES
  ('clando', 150, 100, 250),   -- valeurs FICTIVES, à remplacer par le relevé des prix réels
  ('rakcha', 100,  75, 200),
  ('taxi',   300, 150, 500);

WITH l AS (
  INSERT INTO landmarks (name, quartier, location) VALUES
    ('Grand marche',           'Centre', ST_GeogFromText('POINT(20.8320 13.8290)')),
    ('Marche a betail',        'Est',    ST_GeogFromText('POINT(20.8480 13.8250)')),
    ('Hopital provincial',     'Centre', ST_GeogFromText('POINT(20.8270 13.8330)')),
    ('Universite Adam Barka',  'Nord',   ST_GeogFromText('POINT(20.8300 13.8520)')),
    ('Aeroport',               'Ouest',  ST_GeogFromText('POINT(20.8150 13.8470)')),
    ('Grande mosquee',         'Centre', ST_GeogFromText('POINT(20.8335 13.8300)')),
    ('Gare routiere',          'Sud',    ST_GeogFromText('POINT(20.8350 13.8150)')),
    ('Marche du soir',         'Ouest',  ST_GeogFromText('POINT(20.8200 13.8280)'))
  RETURNING id, name
)
INSERT INTO landmark_names (landmark_id, norm)
SELECT id, lower(name) FROM l;

INSERT INTO landmark_names (landmark_id, norm)
SELECT id, a FROM landmarks, unnest(CASE name
    WHEN 'Grand marche'          THEN ARRAY['souk kabir', 'marche central']
    WHEN 'Hopital provincial'    THEN ARRAY['hopital', 'mustashfa']
    WHEN 'Universite Adam Barka' THEN ARRAY['universite', 'jamia']
    WHEN 'Gare routiere'         THEN ARRAY['gare', 'mahatta']
    ELSE ARRAY[]::text[] END) AS a;
