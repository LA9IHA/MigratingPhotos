SELECT 'id', 'file', 'date_available', 'date_creation', 'name', 'comment', 'author', 'hit', 'filesize', 'width', 'height', 'coi', 'representative_ext', 'date_metadata_update', 'rating_score', 'path', 'storage_category_id', 'level', 'md5sum', 'added_by', 'rotation', 'latitude', 'longitude', 'lastmodified'
 UNION ALL SELECT * FROM album.images
 INTO OUTFILE 'photos.csv'
 FIELDS TERMINATED BY ',' ENCLOSED BY '\"';

SELECT 'id', 'name', 'id_uppercat', 'comment', 'dir', 'rank', 'status', 'site_id', 'visible', 'representative_picture_id', 'uppercats', 'commentable', 'global_rank', 'image_order', 'permalink', 'lastmodified'
UNION ALL SELECT * FROM categories
INTO OUTFILE 'categories.csv'
FIELDS TERMINATED BY ',' ENCLOSED BY '\"';
