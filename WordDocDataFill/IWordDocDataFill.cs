using System.Collections.Generic;

namespace WordDocDataFill
{
    /// <summary>
    /// Fill Word document
    /// </summary>
    public interface IWordDocDataFill
    {
        /// <summary>
        /// Fill Word document data from dictionary
        /// </summary>
        /// <param name="values">Dictionary with key-value collection. 
        /// Key - element in the document to be found and replaced. For example "[LastName]". 
        /// Value - the value which is replaced by the element specified in the key.</param>
        /// <param name="exitDocName">Exit file name</param>
        void FillDocument(Dictionary<string, string> values, string exitDocName);
    }
}
