SELECT
  SocAdhNum,
  SocAdhSocNum,
  ExeExercice,
  ExeLibelle,
  SocAdhDate,
  SocAdhEtat
FROM
  TBLEXERCICES
  INNER JOIN TBLSOCIETESADHERENTS ON TBLEXERCICES.ExeNum = TBLSOCIETESADHERENTS.SocAdhExeNum
ORDER BY
  SocAdhSocNum,
  SocAdhDate DESC,
  ExeLibelle DESC;
