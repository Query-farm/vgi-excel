using System;
using System.Collections.Generic;
using System.Globalization;
using System.Numerics;
using System.Runtime.InteropServices;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace QueryFarm.Vgi.ExcelDna;

// Read materialized chunks, not deprecated C result cells (which truncate NULs
// and nanosecond timestamps). All vector pointers remain owned by their chunk.
internal sealed partial class NativeHaybarnApi
{
    [UnmanagedFunctionPointer(CallingConvention.Cdecl)] private delegate IntPtr ChunkDelegate(Result result, ulong index);
    [UnmanagedFunctionPointer(CallingConvention.Cdecl)] private delegate IntPtr IndexedPointer(IntPtr value, ulong index);
    [UnmanagedFunctionPointer(CallingConvention.Cdecl)] private delegate IntPtr UnaryPointer(IntPtr value);
    [UnmanagedFunctionPointer(CallingConvention.Cdecl)] private delegate ulong UnaryCount(IntPtr value);
    [UnmanagedFunctionPointer(CallingConvention.Cdecl)] private delegate IntPtr CreateLong(long value);
    [UnmanagedFunctionPointer(CallingConvention.Cdecl)] private delegate IntPtr CreateInt(int value);
    [StructLayout(LayoutKind.Sequential)] private struct Interval { internal int Months, Days; internal long Micros; }
    [UnmanagedFunctionPointer(CallingConvention.Cdecl)] private delegate IntPtr CreateInterval(Interval value);
    private ChunkDelegate GetChunk = null!;
    private IndexedPointer ChunkVector = null!, StructVector = null!, StructType = null!, StructName = null!, EnumName = null!;
    private UnaryPointer VectorData = null!, VectorValidity = null!, ListVector = null!, ListType = null!, ArrayVector = null!, ArrayType = null!, ValueText = null!;
    private UnaryCount ChunkSize = null!, ArraySize = null!, StructCount = null!;
    private TypeDelegate DecimalStorage = null!, EnumStorage = null!;
    private DestroyDelegate DestroyChunk = null!, DestroyValue = null!;
    private CreateInt CreateDate = null!;
    private CreateInterval MakeInterval = null!;
    private UnaryPointer MapKeyType = null!, MapValueType = null!;
    private IndexedPointer UnionType = null!;
    private readonly Dictionary<int, CreateLong> temporal = new();

    private void InitializeChunks()
    {
        MapKeyType = Get<UnaryPointer>("duckdb_map_type_key_type"); MapValueType = Get<UnaryPointer>("duckdb_map_type_value_type"); UnionType = Get<IndexedPointer>("duckdb_union_type_member_type");
        GetChunk = Get<ChunkDelegate>("duckdb_result_get_chunk"); ChunkVector = Get<IndexedPointer>("duckdb_data_chunk_get_vector");
        ChunkSize = Get<UnaryCount>("duckdb_data_chunk_get_size"); DestroyChunk = Get<DestroyDelegate>("duckdb_destroy_data_chunk");
        VectorData = Get<UnaryPointer>("duckdb_vector_get_data"); VectorValidity = Get<UnaryPointer>("duckdb_vector_get_validity");
        ListVector = Get<UnaryPointer>("duckdb_list_vector_get_child"); ListType = Get<UnaryPointer>("duckdb_list_type_child_type");
        ArrayVector = Get<UnaryPointer>("duckdb_array_vector_get_child"); ArrayType = Get<UnaryPointer>("duckdb_array_type_child_type"); ArraySize = Get<UnaryCount>("duckdb_array_type_array_size");
        StructVector = Get<IndexedPointer>("duckdb_struct_vector_get_child"); StructType = Get<IndexedPointer>("duckdb_struct_type_child_type"); StructName = Get<IndexedPointer>("duckdb_struct_type_child_name"); StructCount = Get<UnaryCount>("duckdb_struct_type_child_count");
        EnumStorage = Get<TypeDelegate>("duckdb_enum_internal_type"); EnumName = Get<IndexedPointer>("duckdb_enum_dictionary_value"); DecimalStorage = Get<TypeDelegate>("duckdb_decimal_internal_type");
        ValueText = Get<UnaryPointer>("duckdb_get_varchar"); DestroyValue = Get<DestroyDelegate>("duckdb_destroy_value");
        CreateDate = Get<CreateInt>("duckdb_create_date"); MakeInterval = Get<CreateInterval>("duckdb_create_interval");
        foreach (var item in new Dictionary<int,string> { [12]="timestamp", [14]="time", [20]="timestamp_s", [21]="timestamp_ms", [22]="timestamp_ns", [30]="time_tz_value", [31]="timestamp_tz", [39]="time_ns" })
            temporal[item.Key] = Get<CreateLong>("duckdb_create_" + item.Value);
    }

    internal object?[][] ReadRows(Result result, int limit, IntPtr[] logicalTypes, TimeZoneInfo timeZone)
    {
        var rows = new object?[limit][];
        var offset = 0;
        for (ulong chunkIndex = 0; offset < limit; chunkIndex++)
        {
            var chunk = GetChunk(result, chunkIndex);
            if (chunk == IntPtr.Zero) throw new InvalidOperationException("The native result chunk is unavailable.");
            try
            {
                var size = Math.Min(checked((int)ChunkSize(chunk)), limit - offset);
                if (size == 0) throw new InvalidOperationException("The native result chunk is empty.");
                var vectors = new IntPtr[logicalTypes.Length];
                for (var c = 0; c < vectors.Length; c++) vectors[c] = ChunkVector(chunk, (ulong)c);
                for (var r = 0; r < size; r++)
                {
                    var row = rows[offset + r] = new object?[vectors.Length];
                    for (var c = 0; c < vectors.Length; c++)
                    {
                        var value = ReadCell(vectors[c], logicalTypes[c], r, timeZone);
                        row[c] = value is JToken json ? json.ToString(Formatting.None) : value;
                    }
                }
                offset += size;
            }
            finally { DestroyChunk(ref chunk); }
        }
        return rows;
    }

    private object? ReadCell(IntPtr vector, IntPtr logical, int row, TimeZoneInfo timeZone)
    {
        var validity = VectorValidity(vector);
        if (validity != IntPtr.Zero && ((ulong)Marshal.ReadInt64(validity, row / 64 * 8) & (1UL << (row % 64))) == 0) return null;
        var data = VectorData(vector);
        var type = TypeId(logical);
        if (type >= 2 && type <= 9 || type == 16 || type == 32)
            return NativeHaybarnSession.Cell(Integer(data, row, type).ToString(CultureInfo.InvariantCulture), type);
        if (type == 1) return Marshal.ReadByte(data, row) != 0;
        if (type == 10 || type == 11)
        {
            double value;
            if (type == 10) { var values = new float[1]; Marshal.Copy(IntPtr.Add(data,row*4), values,0,1); value=values[0]; }
            else { var values = new double[1]; Marshal.Copy(IntPtr.Add(data,row*8),values,0,1); value=values[0]; }
            return double.IsNaN(value) ? (object)"NaN" : double.IsPositiveInfinity(value) ? "Infinity" : double.IsNegativeInfinity(value) ? "-Infinity" : value == 0 ? 0d : value;
        }
        if (type == 19)
        {
            var number = Integer(data,row,DecimalStorage(logical));
            var scale = DecimalScale(logical);
            var text = BigInteger.Abs(number).ToString(CultureInfo.InvariantCulture).PadLeft(scale+1,'0');
            if (scale > 0) text=text.Insert(text.Length-scale,".");
            if (number.Sign < 0) text="-"+text;
            return NativeHaybarnSession.Cell(text,type);
        }
        if (type == 17) { var bytes = StringBytes(data,row); return System.Text.Encoding.UTF8.GetString(bytes); }
        if (type == 18)
        {
            var bytes=StringBytes(data,row); var text=new System.Text.StringBuilder();
            foreach(var b in bytes) { if(b>=32 && b<=126 && b!=92) text.Append((char)b); else text.Append("\\x"+b.ToString("X2",CultureInfo.InvariantCulture)); }
            return text.ToString();
        }
        if (type == 13) return FormatValue(CreateDate(Marshal.ReadInt32(data,row*4)));
        if (temporal.TryGetValue(type,out var create))
        {
            var text = FormatValue(create(Marshal.ReadInt64(data,row*8)));
            if (type == 31 && DateTimeOffset.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.None, out var timestamp))
                return TimeZoneInfo.ConvertTime(timestamp,timeZone).ToString("yyyy-MM-dd HH:mm:ss.FFFFFFzzz",CultureInfo.InvariantCulture);
            return text;
        }
        if (type == 15) return FormatValue(MakeInterval(Marshal.PtrToStructure<Interval>(IntPtr.Add(data,row*16))));
        if (type == 23)
        {
            var name=EnumName(logical,(ulong)Integer(data,row,EnumStorage(logical)));
            try {return String(name);} finally {Free(name);}
        }
        if (type == 24 || type == 33)
        {
            var childType=type==24 ? ListType(logical) : ArrayType(logical);
            try
            {
                var count=type==24 ? checked((int)Marshal.ReadInt64(data,row*16+8)) : checked((int)ArraySize(logical));
                var start=type==24 ? checked((int)Marshal.ReadInt64(data,row*16)) : checked(row*count);
                var child=type==24 ? ListVector(vector) : ArrayVector(vector);
                var array=new JArray(); for(var i=0;i<count;i++) array.Add(Token(ReadCell(child,childType,start+i,timeZone))); return array;
            }
            finally {DestroyLogicalType(ref childType);}
        }
        if(type==25)
        {
            var value=new JObject();
            for(ulong i=0;i<StructCount(logical);i++)
            {
                var childType=StructType(logical,i);var name=StructName(logical,i);
                try {value[String(name)]=Token(ReadCell(StructVector(vector,i),childType,row,timeZone));}
                finally {Free(name);DestroyLogicalType(ref childType);}
            }
            return value;
        }
        if(type==27)
        {
            var bytes=new byte[17];Marshal.Copy(IntPtr.Add(data,row*16),bytes,0,16);bytes[15]^=0x80;
            var hex=new BigInteger(bytes).ToString("x",CultureInfo.InvariantCulture).PadLeft(32,'0');if(hex.Length>32)hex=hex.Substring(hex.Length-32);
            return hex.Substring(0,8)+"-"+hex.Substring(8,4)+"-"+hex.Substring(12,4)+"-"+hex.Substring(16,4)+"-"+hex.Substring(20);
        }
        if(type==26)
        {
            var child=ListVector(vector);var keys=StructVector(child,0);var values=StructVector(child,1);
            var keyType=MapKeyType(logical);var valueType=MapValueType(logical);
            try
            {
                var start=checked((int)Marshal.ReadInt64(data,row*16));var count=checked((int)Marshal.ReadInt64(data,row*16+8));var map=new JObject();
                for(var i=0;i<count;i++) map[Convert.ToString(ReadCell(keys,keyType,start+i,timeZone),CultureInfo.InvariantCulture)??""]=Token(ReadCell(values,valueType,start+i,timeZone));
                return map;
            }
            finally {DestroyLogicalType(ref keyType);DestroyLogicalType(ref valueType);}
        }
        if(type==28)
        {
            var tag=Marshal.ReadByte(VectorData(StructVector(vector,0)),row);var childType=UnionType(logical,tag);
            try{return ReadCell(StructVector(vector,(ulong)tag+1),childType,row,timeZone);}finally{DestroyLogicalType(ref childType);}
        }
        if(type==29)
        {
            var bytes=StringBytes(data,row);var text=new System.Text.StringBuilder();
            for(var bit=bytes[0];bit<(bytes.Length-1)*8;bit++) text.Append((bytes[1+bit/8] & (1<<(7-bit%8)))!=0?'1':'0');
            return text.ToString();
        }
        if(type==35)
        {
            var bytes=StringBytes(data,row);var negative=(bytes[0]&0x80)==0;var magnitude=new byte[bytes.Length-2];
            for(var i=3;i<bytes.Length;i++) magnitude[bytes.Length-1-i]=negative?(byte)~bytes[i]:bytes[i];
            var value=new BigInteger(magnitude);if(negative)value=-value;
            return value.ToString(CultureInfo.InvariantCulture);
        }
        if(type==36) return null;
        throw new NotSupportedException("This native result type is not supported yet. Cast the column to VARCHAR in SQL.");
    }
    private static JToken Token(object? value) => value is null ? JValue.CreateNull() : value is JToken token ? token : JToken.FromObject(value);
    private static BigInteger Integer(IntPtr data,int row,int type)
    {
        switch(type)
        {
            case 2:return unchecked((sbyte)Marshal.ReadByte(data,row));case 3:return Marshal.ReadInt16(data,row*2);case 4:return Marshal.ReadInt32(data,row*4);case 5:return Marshal.ReadInt64(data,row*8);
            case 6:return Marshal.ReadByte(data,row);case 7:return unchecked((ushort)Marshal.ReadInt16(data,row*2));case 8:return unchecked((uint)Marshal.ReadInt32(data,row*4));case 9:return unchecked((ulong)Marshal.ReadInt64(data,row*8));
            default:var bytes=new byte[type==32?17:16];Marshal.Copy(IntPtr.Add(data,row*16),bytes,0,16);return new BigInteger(bytes);
        }
    }
    private static byte[] StringBytes(IntPtr data,int row)
    {
        var field=IntPtr.Add(data,checked(row*16));var length=Marshal.ReadInt32(field);var pointer=length<=12?IntPtr.Add(field,4):Marshal.ReadIntPtr(field,8);
        var bytes=new byte[length];Marshal.Copy(pointer,bytes,0,length);return bytes;
    }
    private string FormatValue(IntPtr value)
    {
        try {var text=ValueText(value);try{return String(text);}finally{Free(text);}}
        finally {DestroyValue(ref value);}
    }
}
