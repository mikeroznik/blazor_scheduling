using OfficeOpenXml;

namespace ScheduleExcelReader.Data
{
    public class ScheduleService
    {
        public List<Game> GetSchedule(List<ScoreKeeper> scoreKeeperList)
        {
            List<Game> schedule = new List<Game>();
            string filePath = "C:/Users/Igor/Desktop/23-24ScorekeeperSchedule.xlsx";

            FileInfo fileInfo = new FileInfo(filePath);

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using(ExcelPackage excelPackage = new ExcelPackage(fileInfo)) 
            {
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets["Sheet1"];
                int totalColumn = worksheet.Dimension.End.Column;
                int totalRow = worksheet.Dimension.End.Row;

                for (int row = 7; row <= totalRow; row++)
                {
                    Game game = new Game();
                    for (int col = 1; col <= totalColumn; col++)
                    {
                        if (col == 6) game.GameDateAndTime = 
                                Convert.ToDateTime(
                                    String.Concat(
                                        DateTime.FromOADate(double.Parse(worksheet.Cells[row, col-1].Value.ToString())).Date.ToShortDateString(), 
                                        " ", 
                                        Convert.ToDateTime(worksheet.Cells[row, col].Value).ToShortTimeString())
                                    );
						if (col == 8) game.Location = (worksheet.Cells[row, col - 1].Value == null) ?
								worksheet.Cells[row, col].Value.ToString() :
								String.Concat(worksheet.Cells[row, col].Value.ToString(), " Rink ", worksheet.Cells[row, col - 1].Value.ToString());
                        if (col == 10) game.AwayTeam = worksheet.Cells[row, col].Value.ToString();
                        if (col == 11) game.HomeTeam = worksheet.Cells[row, col].Value.ToString();
                        if (col == 12) game.ScoreKeeper = 
                                worksheet.Cells[row, col].Value != null  ? 
                                    scoreKeeperList.Where(sk => sk.Name == worksheet.Cells[row, col].Value.ToString()).First() : new ScoreKeeper();
                    }
                    schedule.Add(game);
                }
            }

            return schedule;
        }

		public List<ScoreKeeper> GetScorekeepers()
		{
			List<ScoreKeeper> scoreKeepers = new List<ScoreKeeper>();
			string filePath = "C:/Users/Igor/Desktop/23-24ScorekeeperSchedule.xlsx";

			FileInfo fileInfo = new FileInfo(filePath);

			ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
			using (ExcelPackage excelPackage = new ExcelPackage(fileInfo))
			{
				ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets["ScoreKeeper"];
				int totalColumn = worksheet.Dimension.End.Column;
				int totalRow = worksheet.Dimension.End.Row;

				for (int row = 2; row <= totalRow; row++)
				{
					ScoreKeeper scoreKeeper = new ScoreKeeper();
					for (int col = 1; col <= totalColumn; col++)
					{
						if (col == 1) scoreKeeper.Name = worksheet.Cells[row, col].Value.ToString();		
                        if (col == 2) scoreKeeper.NoNoTeams = worksheet.Cells[row, col].Value?.ToString().Split(",").ToList();
						if (col == 4) scoreKeeper.OnlyRinks = worksheet.Cells[row, col].Value?.ToString().Split(",").ToList();
					}
                    scoreKeepers.Add(scoreKeeper);
				}
			}

			return scoreKeepers;
		}
	}
}
