SELECT * FROM imdb_final
ORDER BY moviename
SELECT * FROM tomato_final
SELECT * FROM meta_final

ALTER TABLE imdb_final
ALTER COLUMN ratings float
ALTER TABLE tomato_final
ALTER COLUMN rating float
ALTER TABLE meta_final
ALTER COLUMN rating float

SELECT MovieName, IRatings, NewTRatings, NewMRatings,
CAST((IRatings + CAST((NewTRatings + IRatings)/2 AS Float) + CAST((NewMRatings + IRatings)/2 AS Float)) / 3 AS Float) AS FinalRatings
FROM (
SELECT MovieName1 As MovieName, IRatings, 
CASE 
WHEN MRatings is NULL THEN CAST((TRatings + IRatings)/2 AS Float)
ELSE MRatings 
END AS NewMRatings,
CASE 
WHEN TRatings is NULL THEN CAST((MRatings + IRatings)/2 AS Float)
ELSE TRatings
END AS NewTRatings
FROM (
SELECT imdb_final.moviename AS MovieName1, tomato_final.moviename AS MovieName2, meta_final.moviename AS MovieName3, meta_final.rating AS MRatings, 
tomato_final.rating AS TRatings, imdb_final.ratings AS IRatings
FROM imdb_final
FULL JOIN tomato_final ON UPPER(imdb_final.moviename) = UPPER(tomato_final.moviename)
FULL JOIN meta_final ON UPPER(imdb_final.moviename) = UPPER(meta_final.moviename)
WHERE NOT(meta_final.rating IS NULL AND tomato_final.rating IS NULL) 
AND NOT(tomato_final.rating IS NULL AND imdb_final.ratings IS NULL) 
AND NOT(imdb_final.ratings IS NULL AND meta_final.rating IS NULL)
GROUP BY tomato_final.moviename, imdb_final.moviename, meta_final.moviename, meta_final.rating, tomato_final.rating, imdb_final.ratings) AS T1) AS T2
ORDER BY CAST((IRatings + CAST((NewTRatings + IRatings)/2 AS Float) + CAST((NewMRatings + IRatings)/2 AS Float)) / 3 AS Float) DESC

