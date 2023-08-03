using System;

namespace LibSrd
{
    /// <summary>
    /// The result class combines a value result (of any type) with a success indication and and error message.
    /// It consists of a class named Result and a generic class of Result\<T\>.
    /// To set success: : return Result.Ok(value);
    /// To set fail:      return Result.Fail\<type\>("Error message").
    /// 
    /// Posted by Achraf Chennan July 2020
    /// https://achraf-chennan.medium.com/using-the-result-class-in-c-519da90351f0
    /// </summary>
    public class Result
    {
        /// <summary>Is true if Result is successful</summary>
        public bool Success { get; private set; }

        /// <summary>The Error message</summary>
        public string Error { get; private set; }

        /// <summary>Is true if Result is a failure (simply the inverse of Success)</summary>
        public bool IsFailure { get { return !Success; } }
        //public bool IsFailure => !Success;

        /// <summary>
        /// Dont normally directly use
        /// </summary>
        /// <param name="success"></param>
        /// <param name="error"></param>
        protected Result(bool success, string error)
        {
            if (success && error != string.Empty)
                throw new InvalidOperationException();
            if (!success && error == string.Empty)
                throw new InvalidOperationException();
            Success = success;
            Error = error;
        }

        /// <summary>
        /// Use this method to return a Fail result where there would have been no associated value
        /// </summary>
        /// <param name="message"></param>
        /// <returns></returns>
        public static Result Fail(string message)
        {
            return new Result(false, message);
        }

        /// <summary>
        /// Use this method to return a fail result.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="message"></param>
        /// <returns></returns>
        public static Result<T> Fail<T>(string message)
        {
            return new Result<T>(default(T), false, message);
        }

        /// <summary>
        /// Use this method to return a fail result and an error value
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="message"></param>
        /// <param name="value"></param>
        /// <returns></returns>
        public static Result<T> Fail<T>(string message, T value)
        {
            return new Result<T>(value, false, message);
        }

        /// <summary>
        /// Use this method to return a successful result with no associated value
        /// </summary>
        /// <returns></returns>
        public static Result Ok()
        {
            return new Result(true, string.Empty);
        }

        /// <summary>
        /// Use this method to return a successful result with the returned generic value
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="value"></param>
        /// <returns></returns>
        public static Result<T> Ok<T>(T value)
        {
            return new Result<T>(value, true, string.Empty);
        }
    }

    /// <summary>
    /// Dont normally directly use
    /// </summary>
    public class Result<T> : Result
    {
        protected internal Result(T value, bool success, string error) : base(success, error)
        {
            Value = value;
        }

        public T Value { get; set; }
    }
}
