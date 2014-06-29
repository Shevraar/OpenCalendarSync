using System;

namespace Acco.Calendar.Person
{
    public interface IPerson
    {
        string Name { get; set; }
        string FirstName { get; set; }
        string LastName { get; set; }
    }

	public class GenericPerson : IPerson
	{
        public string Name { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }
	}
}