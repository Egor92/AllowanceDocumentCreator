using System.Diagnostics;

namespace AllowanceDocumentCreator
{
    [DebuggerDisplay("IsSuccess: {" + nameof(IsSuccess) + "}; Message: {" + nameof(Message) + "}")]
    public class Result
    {
        protected Result(bool isSuccess, string message = null)
        {
            IsSuccess = isSuccess;
            Message = message;
        }

        public bool IsSuccess { get; }

        public string Message { get;  }

        public static Result Success()
        {
            return new Result(true);
        }

        public static Result<T> Success<T>(T data)
        {
            return new Result<T>(true)
            {
                Data = data,
            };
        }

        public static Result Fault(string message)
        {
            return new Result(false, message);
        }

        public static Result<T> Fault<T>(string message)
        {
            return new Result<T>(false, message);
        }
    }

    [DebuggerDisplay("IsSuccess: {" + nameof(IsSuccess) + "}; Message: {" + nameof(Message) + "}; Data: {Data != null ? Data.GetType() : null}")]
    public sealed class Result<T> : Result
    {
        internal Result(bool isSuccess, string message = null)
            : base(isSuccess, message)
        {
        }

        public T Data { get; set; }
    }
}
