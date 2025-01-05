namespace Core.Domain.Exceptions
{
    public class OfficeOperationException : Exception
    {
        public OfficeOperationException(string message, Exception innerException)
    : base(message, innerException) { }
    }

    public class MappingException : Exception
    {
        public MappingException(string message)
            : base(message) { }
    }
}
