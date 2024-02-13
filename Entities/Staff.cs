namespace Excel2Word.Entities
{
    public class Staff
    {
        public string LastName { get; set; }
        public string FirstName { get; set; }
        public string Middlename { get; set; }
        public int CountProblem { get; set; }

        public override string ToString()
        {
            var result = LastName;
            result += FirstName.Length > 0 ? " " + FirstName[0] + "." : "";
            result += Middlename.Length > 0 ? " " + Middlename[0] + "." : "";
            return result;
        }
    }
}