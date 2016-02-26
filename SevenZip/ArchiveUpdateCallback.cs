/*  This file is part of SevenZipSharp.

    SevenZipSharp is free software: you can redistribute it and/or modify
    it under the terms of the GNU Lesser General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.

    SevenZipSharp is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU Lesser General Public License for more details.

    You should have received a copy of the GNU Lesser General Public License
    along with SevenZipSharp.  If not, see <http://www.gnu.org/licenses/>.
*/

using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
#if MONO
using SevenZip.Mono.COM;
#endif

namespace SevenZip
{
#if UNMANAGED
#if COMPRESS
    /// <summary>
    /// Archive update callback to handle the process of packing files
    /// </summary>
    internal sealed class ArchiveUpdateCallback : CallbackBase, IArchiveUpdateCallback, ICryptoGetTextPassword2,
                                                  IDisposable
    {
        #region Fields
        /// <summary>
        /// _files.Count if do not count directories
        /// </summary>
        private int _actualFilesCount;

        /// <summary>
        /// For Compressing event.
        /// </summary>
        private long _bytesCount;

        private long _bytesWritten;
        private long _bytesWrittenOld;
        private SevenZipCompressor _compressor;

        /// <summary>
        /// No directories.
        /// </summary>
        private bool _directoryStructure;

        /// <summary>
        /// Rate of the done work from [0, 1]
        /// </summary>
        private float _doneRate;

        /// <summary>
        /// The names of the archive entries
        /// </summary>
        private string[] _entries;

        /// <summary>
        /// Array of files to pack
        /// </summary>
        private FileInfo[] _files;
        private string[] _altNames;

        private InStreamWrapper _fileStream;

        private uint _indexInArchive;
        private uint _indexOffset;

        /// <summary>
        /// Common root of file names length.
        /// </summary>
        private int _rootLength;

        /// <summary>
        /// Input streams to be compressed.
        /// </summary>
        private Stream[] _streams;

        private FileInfo[] _streamFileInfos;

        private UpdateData _updateData;
        private List<InStreamWrapper> _wrappersToDispose;

        private readonly object _lockObject = new object();

        /// <summary>
        /// Gets or sets the default item name used in MemoryStream compression.
        /// </summary>
        public string DefaultItemName { private get; set; }

        /// <summary>
        /// Gets or sets the value indicating whether to compress as fast as possible, without calling events.
        /// </summary>
        public bool FastCompression { private get; set; } 
#if !WINCE
        private int _memoryPressure;
#endif
        #endregion

        #region Constructors

        /// <summary>   Initializes a new instance of the ArchiveUpdateCallback class.</summary>
        /// <param name="files">                Array of files to pack.</param>
        /// <param name="altNames">             List of names of the alternates.</param>
        /// <param name="rootLength">           Common file names root length.</param>
        /// <param name="compressor">           The owner of the callback.</param>
        /// <param name="updateData">           The compression parameters.</param>
        /// <param name="directoryStructure">   Preserve directory structure.</param>
        public ArchiveUpdateCallback(
            FileInfo[] files, string[] altNames, int rootLength,
            SevenZipCompressor compressor, UpdateData updateData, bool directoryStructure)
        {
            Init(files, altNames, rootLength, compressor, updateData, directoryStructure);
        }

        /// <summary>   Initializes a new instance of the ArchiveUpdateCallback class.</summary>
        /// <param name="files">                Array of files to pack.</param>
        /// <param name="altNames">             List of names of the alternates.</param>
        /// <param name="rootLength">           Common file names root length.</param>
        /// <param name="password">             The archive password.</param>
        /// <param name="compressor">           The owner of the callback.</param>
        /// <param name="updateData">           The compression parameters.</param>
        /// <param name="directoryStructure">   Preserve directory structure.</param>
        public ArchiveUpdateCallback(
            FileInfo[] files, string[] altNames, int rootLength, string password,
            SevenZipCompressor compressor, UpdateData updateData, bool directoryStructure)
            : base(password)
        {
            Init(files, altNames, rootLength, compressor, updateData, directoryStructure);
        }

        /// <summary>
        /// Initializes a new instance of the ArchiveUpdateCallback class
        /// </summary>
        /// <param name="stream">The input stream</param>
        /// <param name="compressor">The owner of the callback</param>
        /// <param name="updateData">The compression parameters.</param>
        /// <param name="directoryStructure">Preserve directory structure.</param>
        public ArchiveUpdateCallback(
            Stream stream, SevenZipCompressor compressor, UpdateData updateData, bool directoryStructure)
        {
            Init(stream, compressor, updateData, directoryStructure);
        }

        /// <summary>
        /// Initializes a new instance of the ArchiveUpdateCallback class
        /// </summary>
        /// <param name="stream">The input stream</param>
        /// <param name="password">The archive password</param>
        /// <param name="compressor">The owner of the callback</param>
        /// <param name="updateData">The compression parameters.</param>
        /// <param name="directoryStructure">Preserve directory structure.</param>
        public ArchiveUpdateCallback(
            Stream stream, string password, SevenZipCompressor compressor, UpdateData updateData,
            bool directoryStructure)
            : base(password)
        {
            Init(stream, compressor, updateData, directoryStructure);
        }

        /// <summary>
        /// Initializes a new instance of the ArchiveUpdateCallback class
        /// </summary>
        /// <param name="streamDict">Dictionary&lt;file stream, name of the archive entry&gt;</param>
        /// <param name="compressor">The owner of the callback</param>
        /// <param name="updateData">The compression parameters.</param>
        /// <param name="directoryStructure">Preserve directory structure.</param>
        public ArchiveUpdateCallback(
            Dictionary<string, Stream> streamDict,
            SevenZipCompressor compressor, UpdateData updateData, bool directoryStructure)
        {
            Init(streamDict, compressor, updateData, directoryStructure);
        }

        /// <summary>
        /// Initializes a new instance of the ArchiveUpdateCallback class
        /// </summary>
        /// <param name="streamDict">Dictionary&lt;file stream, name of the archive entry&gt;</param>
        /// <param name="password">The archive password</param>
        /// <param name="compressor">The owner of the callback</param>
        /// <param name="updateData">The compression parameters.</param>
        /// <param name="directoryStructure">Preserve directory structure.</param>
        public ArchiveUpdateCallback(
            Dictionary<string, Stream> streamDict, string password,
            SevenZipCompressor compressor, UpdateData updateData, bool directoryStructure)
            : base(password)
        {
            Init(streamDict, compressor, updateData, directoryStructure);
        }

        private void CommonInit(SevenZipCompressor compressor, UpdateData updateData, bool directoryStructure)
        {
            _compressor = compressor;
            _indexInArchive = updateData.FilesCount;
            _indexOffset = updateData.Mode != InternalCompressionMode.Append ? 0 : _indexInArchive;
            if (_compressor.ArchiveFormat == OutArchiveFormat.Zip)
            {
                _wrappersToDispose = new List<InStreamWrapper>();
            }
            _updateData = updateData;
            _directoryStructure = directoryStructure;
            DefaultItemName = "default";            
        }

        private void Init(
            FileInfo[] files, string[] altNames, int rootLength, SevenZipCompressor compressor,
            UpdateData updateData, bool directoryStructure)
        {
            _files = files;
            _rootLength = rootLength;
            _altNames = altNames;

            if (files != null)
            {
                foreach (var fi in files)
                {
                    if (fi.Exists)
                    {
                        if ((fi.Attributes & FileAttributes.Directory) == 0)
                        {
                            _bytesCount += fi.Length;
                            _actualFilesCount++;
                        }
                    }
                }
            }
            CommonInit(compressor, updateData, directoryStructure);
        }

        private void Init(
            Stream stream, SevenZipCompressor compressor, UpdateData updateData, bool directoryStructure)
        {
            _fileStream = new InStreamWrapper(stream, false);
            _fileStream.BytesRead += IntEventArgsHandler;
            _actualFilesCount = 1;
            try
            {
                _bytesCount = stream.Length;
            }
            catch (NotSupportedException)
            {
                _bytesCount = -1;
            }
            try
            {
                stream.Seek(0, SeekOrigin.Begin);
            }
            catch (NotSupportedException)
            {
                _bytesCount = -1;
            }
            CommonInit(compressor, updateData, directoryStructure);
        }

        private void Init(
            Dictionary<string, Stream> streamDict,
            SevenZipCompressor compressor, UpdateData updateData, bool directoryStructure)
        {
            _streams = new Stream[streamDict.Count];
            _streamFileInfos = new FileInfo[streamDict.Count];
            streamDict.Values.CopyTo(_streams, 0);
            _entries = new string[streamDict.Count];
            streamDict.Keys.CopyTo(_entries, 0);
            _actualFilesCount = streamDict.Count;
            for(int i = 0 ; i < _streams.Length; i++)
            {
                var str =_streams[i];
                if (str != null)
                    _bytesCount += str.Length;
                var fileStream = str as FileStream;
                if (fileStream != null)
                {
                    try
                    {
                        var fileInfo = new FileInfo(fileStream.Name);
                        if (fileInfo.Exists)
                            _streamFileInfos[i] = fileInfo;
                    }
                    catch
                    {
                        //ignored.
                    }
                }
            }
            CommonInit(compressor, updateData, directoryStructure);
        }

        #endregion

        /// <summary>
        /// Gets or sets the dictionary size.
        /// </summary>
        public float DictionarySize
        {
            set
            {
#if !WINCE
                _memoryPressure = (int)(value * 1024 * 1024);
                GC.AddMemoryPressure(_memoryPressure);
#endif
            }
        }

        /// <summary>
        /// Raises events for the GetStream method.
        /// </summary>
        /// <param name="index">The current item index.</param>
        /// <returns>True if not cancelled; otherwise, false.</returns>
        private bool EventsForGetStream(uint index)
        {
            if (!FastCompression)
            {
                if (_fileStream != null)
                {
                    _fileStream.BytesRead += IntEventArgsHandler;
                }
                _doneRate += 1.0f / _actualFilesCount;
                var fiea = new FileNameEventArgs(_files != null? _files[index].Name : _entries[index],
                                                 PercentDoneEventArgs.ProducePercentDone(_doneRate));
                OnFileCompression(fiea);
                if (fiea.Cancel)
                {
                    Canceled = true;
                    return false;
                }
            }
            return true;
        }

        #region Events

        /// <summary>
        /// Occurs when the next file is going to be packed.
        /// </summary>
        /// <remarks>Occurs when 7-zip engine requests for an input stream for the next file to pack it</remarks>
        public event EventHandler<FileNameEventArgs> FileCompressionStarted;

        /// <summary>
        /// Occurs when data are being compressed.
        /// </summary>
        public event EventHandler<ProgressEventArgs> Compressing;

        /// <summary>
        /// Occurs when the current file was compressed.
        /// </summary>
        public event EventHandler FileCompressionFinished;

        private void OnFileCompression(FileNameEventArgs e)
        {
            if (FileCompressionStarted != null)
            {
                FileCompressionStarted(this, e);
            }
        }

        private void OnCompressing(ProgressEventArgs e)
        {
            if (Compressing != null)
            {
                Compressing(this, e);
            }
        }

        private void OnFileCompressionFinished(EventArgs e)
        {
            if (FileCompressionFinished != null)
            {
                FileCompressionFinished(this, e);
            }
        }

        #endregion

        #region IArchiveUpdateCallback Members

        public void SetTotal(ulong total) {}

        public void SetCompleted(ref ulong completeValue) {}

        public int GetUpdateItemInfo(uint index, ref int newData, ref int newProperties, ref uint indexInArchive)
        {
            switch (_updateData.Mode)
            {
                case InternalCompressionMode.Create:
                    newData = 1;
                    newProperties = 1;
                    indexInArchive = UInt32.MaxValue;
                    break;
                case InternalCompressionMode.Append:
                    if (index < _indexInArchive)
                    {
                        newData = 0;
                        newProperties = 0;
                        indexInArchive = index;
                    }
                    else
                    {
                        newData = 1;
                        newProperties = 1;
                        indexInArchive = UInt32.MaxValue;
                    }
                    break;
                case InternalCompressionMode.Modify:
                    newData = 0;
                    newProperties = Convert.ToInt32(_updateData.FileNamesToModify.ContainsKey((int)index)
                        && _updateData.FileNamesToModify[(int)index] != null);
                    if (_updateData.FileNamesToModify.ContainsKey((int)index)
                        && _updateData.FileNamesToModify[(int)index] == null)
                    {
                        indexInArchive = (UInt32)_updateData.ArchiveFileData.Count;
                        foreach (KeyValuePair<Int32, string> pairModification in _updateData.FileNamesToModify)
                            if ((pairModification.Key <= index) && (pairModification.Value == null))
                            {
                                do
                                {
                                    indexInArchive--;
                                }
                                while ((indexInArchive > 0) && _updateData.FileNamesToModify.ContainsKey((Int32)indexInArchive)
                                    && (_updateData.FileNamesToModify[(Int32)indexInArchive] == null));
                            }
                    }
                    else
                    {
                        indexInArchive = index;
                    }
                    break;
            }
            return 0;
        }

        public int GetProperty(uint index, ItemPropId propID, ref PropVariant value)
        {
            index -= _indexOffset;
            string val = null;

            FileInfo fileInfo = null;
            if (_files != null && _files.Length > index)
                fileInfo = _files[index];
            else if (_streamFileInfos != null && _streamFileInfos.Length > index)
                fileInfo = _streamFileInfos[index];

            try
            {
                switch (propID)
                {
                    case ItemPropId.IsAnti:
                        value.VarType = VarEnum.VT_BOOL;
                        value.UInt64Value = 0;
                        break;
                    case ItemPropId.SourcePath:
                        if (_updateData.Mode != InternalCompressionMode.Modify)
                        {
                            if (fileInfo != null)
                            {
                                value.VarType = VarEnum.VT_BSTR;
                                val = fileInfo.FullName;
                                value.Value = Marshal.StringToBSTR(val);
                            }
                        }
                        break;
                    case ItemPropId.Path:
                        #region Path

                        value.VarType = VarEnum.VT_BSTR;
                        val = DefaultItemName;
                        if (_updateData.Mode != InternalCompressionMode.Modify)
                        {
                            if (_files == null)
                            {
                                if (_entries != null)
                                {
                                    val = _entries[index];
                                }
                            }
                            else
                            {
                                if (_altNames != null && _altNames[index] != null)
                                {
                                    val = _altNames[index];                                    
                                }
                                else if (_directoryStructure && fileInfo != null)
                                {
                                    if (_rootLength > 0)
                                    {
                                        val = index + fileInfo.FullName.Substring(_rootLength);
                                    }
                                    else
                                    {
                                        val = fileInfo.FullName[0] + fileInfo.FullName.Substring(2);
                                    }
                                }
                                else if (fileInfo != null)
                                {
                                    val = fileInfo.FullName;
                                }
                            }
                        }
                        else
                        {
                            val = _updateData.FileNamesToModify[(int) index];
                        }
                        value.Value = Marshal.StringToBSTR(val);
                        #endregion
                        break;
                    case ItemPropId.IsDirectory:
                        value.VarType = VarEnum.VT_BOOL;
                        if (_updateData.Mode != InternalCompressionMode.Modify)
                        {
                            if (_files == null)
                            {
                                if (_streams == null)
                                {
                                    value.UInt64Value = 0;
                                }
                                else
                                {
                                    value.UInt64Value = (ulong)(_streams[index] == null ? 1 : 0);
                                }
                            }
                            else
                            {
                                value.UInt64Value = (byte)(_files[index].Attributes & FileAttributes.Directory);
                            }
                        }
                        else
                        {
                            value.UInt64Value = Convert.ToUInt64(_updateData.ArchiveFileData[(int) index].IsDirectory);
                        }
                        break;
                    case ItemPropId.Size:
                        #region Size

                        value.VarType = VarEnum.VT_UI8;
                        UInt64 size;
                        if (_updateData.Mode != InternalCompressionMode.Modify)
                        {
                            if (_files == null)
                            {
                                if (_streams == null)
                                {
                                    size = _bytesCount > 0 ? (ulong) _bytesCount : 0;
                                }
                                else if(fileInfo != null)
                                {
                                    size = (ulong)fileInfo.Length;
                                }
                                else
                                {
                                    size = (ulong) (_streams[index] == null? 0 : _streams[index].Length);
                                }
                            }
                            else
                            {
                                size = (fileInfo.Attributes & FileAttributes.Directory) == 0
                                           ? (ulong)fileInfo.Length
                                           : 0;
                            }
                        }
                        else
                        {
                            size = _updateData.ArchiveFileData[(int) index].Size;
                        }
                        value.UInt64Value = size;

                        #endregion
                        break;
                    case ItemPropId.Attributes:
                        value.VarType = VarEnum.VT_UI4;
                        if (_updateData.Mode != InternalCompressionMode.Modify)
                        {
                            if (fileInfo != null)
                            {
                                value.UInt32Value = (uint)fileInfo.Attributes;
                            }
                            else if (_files == null)
                            {
                                if (_streams == null)
                                {
                                    value.UInt32Value = (uint)FileAttributes.Normal;
                                }
                                else
                                {
                                    value.UInt32Value = (uint)(_streams[index] == null ? FileAttributes.Directory : FileAttributes.Normal);
                                }
                            }
                        }
                        else
                        {
                            value.UInt32Value = _updateData.ArchiveFileData[(int) index].Attributes;
                        }
                        break;
                    #region Times
                    case ItemPropId.CreationTime:
                        value.VarType = VarEnum.VT_FILETIME;
                        if (_updateData.Mode != InternalCompressionMode.Modify)
                        {
                            value.Int64Value = fileInfo == null
                                               ? DateTime.Now.ToFileTime()
                                               : fileInfo.CreationTime.ToFileTime();
                        }
                        else
                        {
                            value.Int64Value = _updateData.ArchiveFileData[(int) index].CreationTime.ToFileTime();
                        }
                        break;
                    case ItemPropId.LastAccessTime:
                        value.VarType = VarEnum.VT_FILETIME;
                        if (_updateData.Mode != InternalCompressionMode.Modify)
                        {
                            value.Int64Value = fileInfo == null
                                               ? DateTime.Now.ToFileTime()
                                               : fileInfo.LastAccessTime.ToFileTime();
                        }
                        else
                        {
                            value.Int64Value = _updateData.ArchiveFileData[(int) index].LastAccessTime.ToFileTime();
                        }
                        break;
                    case ItemPropId.LastWriteTime:
                        value.VarType = VarEnum.VT_FILETIME;
                        if (_updateData.Mode != InternalCompressionMode.Modify)
                        {
                            value.Int64Value = fileInfo == null
                                               ? DateTime.Now.ToFileTime()
                                               : fileInfo.LastWriteTime.ToFileTime();
                        }
                        else
                        {
                            value.Int64Value = _updateData.ArchiveFileData[(int) index].LastWriteTime.ToFileTime();
                        }
                        break;
                    #endregion
                    case ItemPropId.Extension:
                        #region Extension

                        value.VarType = VarEnum.VT_BSTR;
                        if (_updateData.Mode != InternalCompressionMode.Modify)
                        {
                            try
                            {
                                val = fileInfo != null
                                      ? fileInfo.Extension.Substring(1)
                                      : _entries == null
                                          ? ""
                                          : Path.GetExtension(_entries[index]);
                                value.Value = Marshal.StringToBSTR(val);
                            }
                            catch (ArgumentException)
                            {
                                value.Value = Marshal.StringToBSTR("");
                            }
                        }
                        else
                        {
                            val = Path.GetExtension(_updateData.ArchiveFileData[(int) index].FileName);
                            value.Value = Marshal.StringToBSTR(val);
                        }

                        #endregion
                        break;
                }
            }
            catch (Exception e)
            {
                AddException(e);
            }
            return 0;
        }

        /// <summary>
        /// Gets the stream for 7-zip library.
        /// </summary>
        /// <param name="index">File index</param>
        /// <param name="inStream">Input file stream</param>
        /// <returns>Zero if Ok</returns>
        public int GetStream(uint index, out 
#if !MONO
		                     ISequentialInStream
#else
		                     HandleRef
#endif
		                     inStream)
        {
            index -= _indexOffset;
            
            if (_files != null)
            {
                _fileStream = null;
                try
                {
                    string fullName = _files[index].FullName;
                    // mmmmm
                    if (File.Exists(fullName))
                    {
                        _fileStream = new InStreamWrapper(
                            new FileStream(fullName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite),
                            true);

                        if (_compressor.ArchiveFormat == OutArchiveFormat.Zip) 
                            _wrappersToDispose.Add(_fileStream);

                    }
                }
                catch (Exception e)
                {
                    AddException(e);
                    inStream = null;
                    return -1;
                }
                inStream = _fileStream;
                if (!EventsForGetStream(index))
                {
                    return -1;
                }
            }
            else
            {
                if (_streams == null)
                {
                    inStream = _fileStream;
                }
                else
                {
                    _fileStream = new InStreamWrapper(_streams[index], true);
                    inStream = _fileStream;
                    if (!EventsForGetStream(index))
                    {
                        return -1;
                    }
                }
            }
            return 0;
        }

        public long EnumProperties(IntPtr enumerator)
        {
            //Not implemented HRESULT
            return 0x80004001L;
        }

        public void SetOperationResult(OperationResult operationResult)
        {
            if (operationResult != OperationResult.Ok && ReportErrors)
            {
                if (_fileStream != null)
                {
                    try
                    {
                        _fileStream.Dispose();
                    }
                    catch (ObjectDisposedException) { }
                    _fileStream = null;
                }

                switch (operationResult)
                {
                    case OperationResult.CrcError:
                        AddException(new ExtractionFailedException("File is corrupted. Crc check has failed."));
                        break;
                    case OperationResult.DataError:
                        AddException(new ExtractionFailedException("File is corrupted. Data error has occured."));
                        break;
                    case OperationResult.UnsupportedMethod:
                        AddException(new ExtractionFailedException("Unsupported method error has occured."));
                        break;
                }
            }
            if (_fileStream != null )
            {
                
                    _fileStream.BytesRead -= IntEventArgsHandler;
                    //Specific Zip implementation - can not Dispose files for Zip.
                    if (_compressor.ArchiveFormat != OutArchiveFormat.Zip) 
                    {
                        try
                        {
                            _fileStream.Dispose();                            
                        }
                        catch (ObjectDisposedException) {}
                    }
                    else
                    {
                        _wrappersToDispose.Add(_fileStream);
                    }                                
                _fileStream = null;
                //remove useless GC.collect
//                GC.Collect();
                // Issue #6987
                //GC.WaitForPendingFinalizers();
            }
            OnFileCompressionFinished(EventArgs.Empty);
        }

        #endregion

        #region ICryptoGetTextPassword2 Members

        public int CryptoGetTextPassword2(ref int passwordIsDefined, out string password)
        {
            passwordIsDefined = String.IsNullOrEmpty(Password) ? 0 : 1;
            password = Password;
            return 0;
        }

        #endregion

        #region IDisposable Members

        public void Dispose()
        {
#if !WINCE
            GC.RemoveMemoryPressure(_memoryPressure);
#endif
            if (_fileStream != null)
            {
                try
                {
                    _fileStream.Dispose();
                }
                catch (ObjectDisposedException) {}
            }
            if (_wrappersToDispose != null)
            {
                foreach (var wrapper in _wrappersToDispose)
                {
                    try
                    {
                        wrapper.Dispose();
                    }
                    catch (ObjectDisposedException) {}
                }
            }
            GC.SuppressFinalize(this);
        }

        #endregion

        private void IntEventArgsHandler(object sender, IntEventArgs e)
        {
            lock (_lockObject)
            {
				//SAB: Fix div zero
                if (_bytesCount == 0)
                    _bytesCount = 1;
                var pold = (byte) ((_bytesWrittenOld*100)/_bytesCount);
                _bytesWritten += e.Value;
                byte pnow;
                if (_bytesCount < _bytesWritten) //this check for ZIP is golden
                {
                    pnow = 100;
                }
                else
                {
                    pnow = (byte)((_bytesWritten * 100) / _bytesCount);
                }
                if (pnow > pold)
                {
                    _bytesWrittenOld = _bytesWritten;
                    OnCompressing(new ProgressEventArgs(pnow, (byte) (pnow - pold)));
                }
            }
        }
    }
#endif
#endif
}
