namespace FORM_TICKETING_SYSTEM.Models
{
    public class Ticket
    {
        public Guid Id { get; set; }
        public string? FirstName { get; set; }
        public string?  MiddleName { get; set; }
        public string? LastName { get; set; }
        public string? Email { get; set; }
        public string? Description { get; set; }
    }
}
