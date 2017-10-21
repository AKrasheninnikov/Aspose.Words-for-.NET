using System;
using System.Collections.Generic;
using System.Linq;

namespace SigningDocumentExample
{
    public class SignPersonList
    {
        public Guid Id { get; internal set; }
        public string Name { get; internal set; }
        public string Position { get; internal set; }
        public byte[] Image { get; internal set; }

        public IEnumerable<SignPersonList> GetAllSignPersons()
        {
            return _signPersons;
        }

        public string GetSignerPositionByName(string signerName)
        {
            return _signPersons.AsEnumerable().Where(p => p.Name == signerName).Select(row => row.Position).FirstOrDefault();
        }

        public Guid GetSignerIdByName(string signerName)
        {
            return _signPersons.AsEnumerable().Where(p => p.Name == signerName).Select(row => row.Id).FirstOrDefault();
        }

        public Byte[] GetSignerImageByName(string signerName)
        {
            return _signPersons.AsEnumerable().Where(p => p.Name == signerName).Select(row => row.Image).FirstOrDefault();
        }

        private readonly List<SignPersonList> _signPersons = SignPersonTestData.SignPersons;
    }

    class SignPersonTestData : SignPersonList
    {
        public static List<SignPersonList> SignPersons = new List<SignPersonList> {
            new SignPersonList { Id = Guid.Parse("CDAA3044-8017-4E07-BFF4-93EA14A3A6C9"), Name = "Hocs", Position = "Head of Corporate Services", Image = ConvertHepler.ConverTestImageToByteArray() },
            new SignPersonList { Id = Guid.Parse("1C22DFF1-B98E-4F65-888F-D55F9A968CD3"), Name = "Dhocs", Position = "Deputy Head of Corporate Services", Image = ConvertHepler.ConverTestImageToByteArray() }
        };
    }
}
