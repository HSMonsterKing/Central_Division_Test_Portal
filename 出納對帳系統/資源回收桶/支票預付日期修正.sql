update �ǲ���� set �w�I��� = �n�����, �n����� = NULL
FROM �ǲ���� INNER JOIN �{���Ƭdï ON �ǲ����.�~ = �{���Ƭdï.�~ AND �ǲ����.�ǲ����X = �{���Ƭdï.�ǲ����X 
WHERE (�{���Ƭdï.���J���B405 > 0 OR �{���Ƭdï.��X���B405 > 0) 
AND NOT (�ǲ����.�פJ�b�� IS NOT NULL AND �ǲ����.�פJ�b�� != '')
AND (�ǲ����.���J���B = 0 OR �ǲ����.���J���B IS NULL)
