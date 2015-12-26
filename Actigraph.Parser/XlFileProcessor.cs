﻿using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace Actigraph.Parser
{
   public class XlFileProcessor
    {
       public IEnumerable<SubjectRecords> LoadFile(string filePath)
       {
           try
           {

          
            List<SubjectRecords> subjectRecordsList = new List<SubjectRecords>();
           var absoluteFilePath = Path.Combine(DirectoryStructure.ActigraphDataFilesFolderName, filePath);
           if (File.Exists(absoluteFilePath))
           {
               var result = LoadXls(absoluteFilePath,"Daily");
               var detailSubjectRecords = (from DataRow dRow in result.Rows
                   select new SubjectData()
                   {
                       ID = dRow["Filename"].ToString().Split('_')[0],
                       Age = Convert.ToInt32(dRow["Age"]),
                       Gender = dRow["Gender"].ToString(),
                       Date = Convert.ToDateTime(dRow["Date"]),
                       DayofWeek = dRow["Day of Week"].ToString(),
                       Axix1Cpm = Convert.ToDouble(dRow["Axis 1 CPM"]),
                       FreedsonBouts = Convert.ToInt64(dRow["Freedson (1998) Bouts"]),
                       Lifestyle = Convert.ToDouble(dRow["Lifestyle"]),
                       Light = Convert.ToDouble(dRow["Light"]),
                       Moderate = Convert.ToDouble(dRow["Moderate"]),
                       Vigorous = Convert.ToDouble(dRow["Vigorous"]),
                       VeryVigorous = Convert.ToDouble(dRow["Very Vigorous"]),
                       Sedentary = Convert.ToDouble(dRow["Sedentary"]),
                       SedentaryBouts = Convert.ToInt64(dRow["Sedentary Bouts"]),
                       StepsCount = Convert.ToInt64(dRow["Steps Counts"]),
                       Time = Convert.ToDouble(dRow["Time"]),
                       VectorMagnitudeCpm = Convert.ToDouble(dRow["Vector Magnitude CPM"])
                   }).OrderBy(c => c.ID);

                var distinctSubjectId = GetDistinctSubjectId(detailSubjectRecords);
               subjectRecordsList.AddRange(from subjectItem in distinctSubjectId
                   let subjectRecords = GetSubjectRecords(detailSubjectRecords, subjectItem)
                   let subjectDatas = (IList<SubjectData>) (subjectRecords as IList<SubjectData> ?? subjectRecords.ToList())
                   let subjectvalidRecords = GetSubjectValidRecords(subjectDatas)
                   let validSubjectRecords = (SubjectData[]) (subjectvalidRecords as SubjectData[] ?? subjectvalidRecords.ToArray())
                   let validSubjectAvgRecords = GetValidDaysAverages(validSubjectRecords, subjectItem)
                   select new SubjectRecords()
                   {
                       SubjectData = subjectDatas, SubjectValidData = validSubjectRecords, SubjectValidAverages = validSubjectAvgRecords, SubjectId = subjectItem
                   });
                
           }
            return subjectRecordsList;
            }
            catch (Exception exp)
            {

                throw exp;
            }
        }
       private DataTable LoadXls(string strFile, String sheetName)
        {
           try
           {
            DataTable dtXls = new DataTable(sheetName);
                var strConnectionString = "";

                if (strFile.Trim().EndsWith(".xlsx"))
                {

                    strConnectionString = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=\"Excel 12.0 Xml;HDR=YES;IMEX=1\";", strFile);

                }
                else if (strFile.Trim().EndsWith(".xls"))
                {

                    strConnectionString = string.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties=\"Excel 8.0;HDR=Yes;IMEX=1\";", strFile);

                }

                OleDbConnection SQLConn = new OleDbConnection(strConnectionString);

                SQLConn.Open();

                OleDbDataAdapter sqlAdapter = new OleDbDataAdapter();

                string sql = "SELECT * FROM [" + sheetName + "$]";

                OleDbCommand selectCmd = new OleDbCommand(sql, SQLConn);

                sqlAdapter.SelectCommand = selectCmd;

                sqlAdapter.Fill(dtXls);

                SQLConn.Close();
         
           return dtXls;
            }
            catch (Exception exp)
            {

                throw exp;
            }

        }

       private IEnumerable<SubjectData> GetSubjectRecords(IEnumerable<SubjectData> subjectData,string subjectId)
       {
           return subjectData.Where(p => p.ID == subjectId);
       }

       private IEnumerable<SubjectData> GetSubjectValidRecords(IEnumerable<SubjectData> subjectData)
       {
            return subjectData.Where(p => p.Time>=600);
        }

       private SubjectValidAverages GetValidDaysAverages(IEnumerable<SubjectData> validSubjectRecords,string subjectId)
       {
           try
           {

            double avgSedentary = 0d;
            double avgLight = 0d;
            double avgLifestyle = 0d;
            double avgModerate = 0d;
            double avgTime = 0d;
            long validDays = 0;
               double movementsPerMinute = 0d;
               double steps=0d;

           var subjectRecords = validSubjectRecords as SubjectData[] ?? validSubjectRecords.ToArray();
            
           validDays = subjectRecords.Count();
            foreach (var rec in subjectRecords)
            {
                avgSedentary += rec.Sedentary;
                avgLifestyle += rec.Lifestyle;
                avgLight += rec.Light;
                avgModerate += rec.Moderate + rec.Vigorous + rec.VeryVigorous;
                avgTime += rec.Time;
                

            }

           return new SubjectValidAverages()
           {
               Id = subjectId,
               AvgLifestyle = Math.Round(avgLifestyle/validDays),
               AvgLight = Math.Round(avgLight /validDays, 2),
               AvgModerate = Math.Round(avgModerate /validDays, 2),
               AvgSedentary = Math.Round(avgSedentary /validDays, 2),
               AvgTime = Math.Round(avgTime /validDays, 2),
               ValidDays = validDays,
               
               

           };
            }
            catch (Exception exp)
            {

                throw exp;
            }
        }

       private IEnumerable<string> GetDistinctSubjectId(IEnumerable<SubjectData> subjectData)
       {
           var distinctRecords = subjectData.Select(p => p.ID).Distinct();
           return distinctRecords;
       }
    }
}
