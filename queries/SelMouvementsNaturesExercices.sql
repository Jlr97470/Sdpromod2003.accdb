SELECT
  MvtSocNum,
  MvtDate,
  MvtNatNum,
  MvtNatCode,
  ExeNum,
  ExeExercice,
  MvtIntitule,
  MvtDebit,
  MvtCredit,
  MvtQuantite,
  MvtNum,
  MvtDossier
FROM
  TBLMOUVEMENTSNATURES
  INNER JOIN (
    TBLEXERCICES
    INNER JOIN TBLMOUVEMENTS ON TBLEXERCICES.ExeNum = TBLMOUVEMENTS.MvtExeNum
  ) ON TBLMOUVEMENTSNATURES.MvtNatNum = TBLMOUVEMENTS.MvtMvtNatNum
ORDER BY
  MvtDate DESC,
  MvtNum DESC;
