namespace Core.Domain.Models
{
    public class DefinedName
    {
        public string Name { get; }
        public string CellReference { get; }

        public DefinedName(string name, string cellReference)
        {
            if (string.IsNullOrWhiteSpace(name))
                throw new ArgumentNullException(nameof(name));

            if (string.IsNullOrWhiteSpace(cellReference))
                throw new ArgumentNullException(nameof(cellReference));

            Name = name;
            CellReference = cellReference;
        }

        public override string ToString()
        {
            return Name;
        }

        public override bool Equals(object? obj)
        {
            if (obj is not DefinedName other)
                return false;

            return Name.Equals(other.Name, StringComparison.OrdinalIgnoreCase);
        }

        public override int GetHashCode()
        {
            return Name.GetHashCode(StringComparison.OrdinalIgnoreCase);
        }
    }
}
