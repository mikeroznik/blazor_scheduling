using OfficeOpenXml;

namespace ScheduleExcelReader.Data
{
    public class ScheduleService
    {
        public List<Game> GetSchedule(List<ScoreKeeper> scoreKeeperList)
        {
            List<Game> schedule = new List<Game>();
            string filePath = "23-24ScorekeeperSchedule.xlsx";

            FileInfo fileInfo = new FileInfo(filePath);

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using(ExcelPackage excelPackage = new ExcelPackage(fileInfo)) 
            {
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets["February"];
                int totalColumn = 7; // worksheet.Dimension.End.Column;
                int totalRow = worksheet.Dimension.End.Row;

                for (int row = 1; row <= totalRow; row++)
                {
                    // Create the games object
                    Game game = new Game();
                    for (int col = 1; col <= totalColumn; col++)
                    {
                        if (col == 3) game.GameDateAndTime = 
                                Convert.ToDateTime(
                                    String.Concat(
                                        DateTime.FromOADate(double.Parse(worksheet.Cells[row, col-1].Value.ToString())).Date.ToShortDateString(), 
                                        " ", 
                                        Convert.ToDateTime(worksheet.Cells[row, col].Value).ToShortTimeString())
                                    );
                        // Locations are now correct on the scoresheet, no more string concat
						if (col == 4) game.Location = String.Concat(worksheet.Cells[row, col].Value.ToString());
                        if (col == 5) game.AwayTeam = worksheet.Cells[row, col].Value.ToString();
                        if (col == 6) game.HomeTeam = worksheet.Cells[row, col].Value.ToString();
                        if (col == 7) game.ScoreKeeper = 
                                worksheet.Cells[row, col].Value != null  ? 
                                    scoreKeeperList.Where(sk => sk.Name == worksheet.Cells[row, col].Value.ToString()).First() : new ScoreKeeper();
                    }

                    // skip the beginner practices
                    if (game.AwayTeam != "No Officials")
                        schedule.Add(game);
                }
            }

            return schedule;
        }

		public List<ScoreKeeper> GetScorekeepers()
		{
			List<ScoreKeeper> scoreKeepers = new List<ScoreKeeper>();
			string filePath = "23-24ScorekeeperSchedule.xlsx";

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
