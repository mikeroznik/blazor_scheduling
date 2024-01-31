namespace ScheduleExcelReader.Data
{
    public class Game
    {
        public DateTime GameDateAndTime { get; set; }
        public string Rink {  get; set; }   
        public string Location { get; set; }
        public string AwayTeam { get; set; }
        public string HomeTeam { get; set; }
        public ScoreKeeper? ScoreKeeper { get; set; }

    }
}
