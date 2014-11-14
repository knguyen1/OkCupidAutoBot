using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OkCupidAutoBot
{
    class ProfileDetails
    {
        private string _profileId;
        private string _userName;
        private int _age;
        private string _location;
        private int _percentage;
        private int _enemy;
        private DateTime _lastOnline;
        private string _orientation;
        private string _ethnicity;
        private string _height;
        private string _bodyType;
        private string _diet;
        private string _smokes;
        private string _drinks;
        private string _drugs;
        private string _religion;
        private string _sign;
        private string _education;
        private string _job;
        private string _income;
        private string _relationshipStatus;
        private string _relationshipType;
        private string _offspring;
        private string _pets;
        private string _speaks;
        private string _createdBy;
        private DateTime _createdDt;

        public string ProfileId { get { return _profileId; } set { _profileId = value; } }
        public string Username { get { return _userName; } set { _userName = value; } }
        public int Age { get { return _age; } set { _age = value; } }
        public string Location { get { return _location; } set { _location = value; } }
        public int Percentage { get { return _percentage; } set { _percentage = value; } }
        public int Enemy { get { return _enemy; } set { _enemy = value; } }
        public DateTime LastOnline { get { return _lastOnline; } set { _lastOnline = value; } }
        public string Orientation { get { return _orientation; } set { _orientation = value; } }
        public string Ethnicity { get { return _ethnicity; } set { _ethnicity = value; } }
        public string Height { get { return _height; } set { _height = value; } }
        public string BodyType { get { return _bodyType; } set { _bodyType = value; } }
        public string Diet { get { return _diet; } set { _diet = value; } }
        public string Smokes { get { return _smokes; } set { _smokes = value; } }
        public string Drinks { get { return _drinks; } set { _drinks = value; } }
        public string Drugs { get { return _drugs; } set { _drugs = value; } }
        public string Religion { get { return _religion; } set { _religion = value; } }
        public string Sign { get { return _sign; } set { _sign = value; } }
        public string Education { get { return _education; } set { _education = value; } }
        public string Job { get { return _job; } set { _job = value; } }
        public string Income { get { return _income; } set { _income = value; } }
        public string RelationshipStatus { get { return _relationshipStatus; } set { _relationshipStatus = value; } }
        public string RelationshipType { get { return _relationshipType; } set { _relationshipType = value; } }
        public string Offspring { get { return _offspring; } set { _offspring = value; } }
        public string Pets { get { return _pets; } set { _pets = value; } }
        public string Speaks { get { return _speaks; } set { _speaks = value; } }
        public string CreatedBy { get { return _createdBy; } set { _createdBy = value; } }
        public DateTime CreatedDt { get { return _createdDt; } set { _createdDt = value; } }
    }
}
