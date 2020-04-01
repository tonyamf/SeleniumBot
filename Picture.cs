namespace SeleniumBot
{
    class Picture
    {
        public string title { get; set; }
        public double rt { get; set; }
        public double vote { get; set; }
        public double place { get; set; }
        public double point { get; set; }
        public double pop { get; set; }
        public Picture(string title, double vote, double place, double rt)
        {
            this.title = title;
            this.rt = rt;
            this.vote = vote;
            this.place = place;

        }

    }
}
