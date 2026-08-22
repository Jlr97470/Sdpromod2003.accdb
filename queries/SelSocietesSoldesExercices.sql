SELECT
  MvtSocNum,
  MvtExeNum,
  ExeLibelle,
  Sum(MvtDebit) AS TotalDebit,
  Sum(MvtCredit) AS TotalCredit,
  Sum([MvtCredit] - [MvtDebit]) AS Solde
FROM
  TBLEXERCICES
  INNER JOIN TBLMOUVEMENTS ON TBLEXERCICES.ExeNum = TBLMOUVEMENTS.MvtExeNum
WHERE
  MvtMvtNatNum <> 4
GROUP BY
  MvtSocNum,
  MvtExeNum,
  ExeLibelle
ORDER BY
  MvtSocNum,
  MvtExeNum;
