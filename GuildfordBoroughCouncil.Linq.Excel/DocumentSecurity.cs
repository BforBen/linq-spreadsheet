namespace GuildfordBoroughCouncil.Security
{
    public static class InformationProtectiveMarking
    {
        public enum Gpms
        {
            NonBusiness,
            Unclassified,
            Protect,
            Restricted
        }

        public enum Gscp
        {
            Official,
            OfficialSensitive,
            Secret,
            TopSecret
        }

        public enum Distribution
        {
            Internal,
            External
        }
    }
}
