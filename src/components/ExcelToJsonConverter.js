import React, { useState, useMemo } from 'react';
import * as XLSX from 'xlsx';

const ExcelToJsonConverter = () => {
    const [data, setData] = useState([]);
    const [loading, setLoading] = useState(false);
    const [currentPage, setCurrentPage] = useState(1);
    const [sortConfig, setSortConfig] = useState({ key: null, direction: 'asc' });
    const [searchTerm, setSearchTerm] = useState("");
    const [editingCell, setEditingCell] = useState(null); // { rowIndex, columnKey }
    const itemsPerPage = 10;

    // --- LOGIC (Filter -> Sort -> Paginate) ---
    const filteredData = useMemo(() => {
        if (!searchTerm) return data;
        return data.filter((row) =>
            Object.values(row).some((val) =>
                String(val).toLowerCase().includes(searchTerm.toLowerCase())
            )
        );
    }, [data, searchTerm]);

    const sortedData = useMemo(() => {
        let sortableItems = [...filteredData];
        if (sortConfig.key !== null) {
            sortableItems.sort((a, b) => {
                const aValue = String(a[sortConfig.key] || "");
                const bValue = String(b[sortConfig.key] || "");
                return sortConfig.direction === 'asc'
                    ? aValue.localeCompare(bValue, undefined, { numeric: true })
                    : bValue.localeCompare(aValue, undefined, { numeric: true });
            });
        }
        return sortableItems;
    }, [filteredData, sortConfig]);

    const displayData = useMemo(() => {
        const start = (currentPage - 1) * itemsPerPage;
        return sortedData.slice(start, start + itemsPerPage);
    }, [sortedData, currentPage]);

    const totalPages = Math.ceil(sortedData.length / itemsPerPage);

    const requestSort = (key) => {
        let direction = 'asc';
        if (sortConfig.key === key && sortConfig.direction === 'asc') {
            direction = 'desc';
        }
        setSortConfig({ key, direction });
    };

    const handleFileUpload = (e) => {
        setLoading(true);
        const file = e.target.files[0];
        if (!file) return;
        const reader = new FileReader();
        reader.onload = (event) => {
            const bstr = event.target.result;
            const workbook = XLSX.read(bstr, { type: 'binary' });
            const workSheet = workbook.Sheets[workbook.SheetNames[0]];
            const json = XLSX.utils.sheet_to_json(workSheet, { header: 0, raw: false, dateNF: "yyyy-mm-dd" });
            setData(json);
            setCurrentPage(1);
            setLoading(false);
        };
        reader.readAsBinaryString(file);
    };

    // --- EDITING LOGIC ---
    const handleCellUpdate = (newValue, rowIndex, columnKey) => {
        const updatedData = [...data];
        // Find the actual index in the master data array
        // because displayData is a sliced/sorted version
        const actualIndex = data.findIndex(row => row === displayData[rowIndex]);

        if (actualIndex > -1) {
            updatedData[actualIndex][columnKey] = newValue;
            setData(updatedData);
        }
        setEditingCell(null);
    };

    // --- EXPORT FUNCTION ---
    const handleExport = () => {
        const worksheet = XLSX.utils.json_to_sheet(sortedData);

        // Calculate column widths
        const objectMaxLength = [];
        sortedData.forEach((row) => {
            Object.keys(row).forEach((key, i) => {
                const value = row[key] ? row[key].toString() : "";
                const width = Math.max(key.length, value.length);
                objectMaxLength[i] = Math.max(objectMaxLength[i] || 0, width);
            });
        });

        // Apply widths (adding 2 for extra padding)
        worksheet["!cols"] = objectMaxLength.map((w) => ({ width: w + 2 }));

        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, "FilteredData");
        XLSX.writeFile(workbook, `Exported_Data.xlsx`);
    };

    // --- STYLES ---
    const styles = {
        container: { padding: '40px', fontFamily: '"Inter", sans-serif', backgroundColor: '#f3f4f6', minHeight: '100vh' },
        card: { backgroundColor: '#fff', borderRadius: '12px', boxShadow: '0 10px 15px -3px rgba(0,0,0,0.1)', padding: '24px' },
        header: { display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '24px' },
        searchInput: { padding: '10px', borderRadius: '8px', border: '1px solid #d1d5db', width: '250px' },
        tableWrapper: { overflowX: 'auto', border: '1px solid #e5e7eb', borderRadius: '8px' },
        table: { width: '100%', borderCollapse: 'collapse', fontSize: '14px' },
        th: { backgroundColor: '#f9fafb', padding: '12px', textAlign: 'left', cursor: 'pointer', borderBottom: '2px solid #e5e7eb' },
        td: { padding: '12px', borderBottom: '1px solid #f3f4f6' },
        btn: { padding: '8px 16px', borderRadius: '6px', border: '1px solid #d1d5db', cursor: 'pointer', backgroundColor: '#fff', fontWeight: '500' },
        exportBtn: { backgroundColor: '#2563eb', color: 'white', border: 'none', padding: '10px 20px', borderRadius: '8px', cursor: 'pointer', fontWeight: '600' }
    };

    return (
        <div style={styles.container}>
            <div style={styles.card}>
                <div style={styles.header}>
                    <h2 style={{ margin: 0 }}>Associate Batch Tracker</h2>
                    <div style={{ display: 'flex', gap: '12px', alignItems: 'center' }}>
                        <input type="file" onChange={handleFileUpload} />
                        {data.length > 0 && (
                            <>
                                <input
                                    type="text"
                                    placeholder="Search..."
                                    value={searchTerm}
                                    onChange={(e) => setSearchTerm(e.target.value)}
                                    style={styles.searchInput}
                                />
                                <button onClick={handleExport} style={styles.exportBtn}>
                                    Download Excel
                                </button>
                            </>
                        )}
                    </div>
                </div>

                {data.length > 0 ? (
                    <>
                        <div style={styles.tableWrapper}>
                            <table style={styles.table}>
                                <thead>
                                    <tr>
                                        <th style={styles.th}>#</th>
                                        {Object.keys(data[0]).map((key) => (
                                            <th key={key} style={styles.th} onClick={() => requestSort(key)}>
                                                {key} {sortConfig.key === key ? (sortConfig.direction === 'asc' ? '▲' : '▼') : '↕'}
                                            </th>
                                        ))}
                                    </tr>
                                </thead>
                                <tbody>
                                    {displayData.map((row, rowIndex) => (
                                        <tr key={rowIndex}>
                                            <td style={styles.td}>{(currentPage - 1) * itemsPerPage + rowIndex + 1}</td>
                                            {Object.keys(row).map((columnKey) => (
                                                <td
                                                    key={columnKey}
                                                    style={styles.td}
                                                    onClick={() => setEditingCell({ rowIndex, columnKey })}
                                                >
                                                    {editingCell?.rowIndex === rowIndex && editingCell?.columnKey === columnKey ? (
                                                        <input
                                                            autoFocus
                                                            style={styles.editInput}
                                                            defaultValue={row[columnKey]}
                                                            onBlur={(e) => handleCellUpdate(e.target.value, rowIndex, columnKey)}
                                                            onKeyDown={(e) => e.key === 'Enter' && handleCellUpdate(e.target.value, rowIndex, columnKey)}
                                                        />
                                                    ) : (
                                                        <span>{row[columnKey]}</span>
                                                    )}
                                                </td>
                                            ))}
                                        </tr>
                                    ))}
                                </tbody>
                            </table>
                        </div>
                        <div style={{ marginTop: '20px', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                            <small style={{ color: '#6b7280' }}>
                                Showing {displayData.length} of {sortedData.length} records
                            </small>
                            <div style={{ display: 'flex', gap: '8px', alignItems: 'center' }}>
                                <button style={styles.btn} onClick={() => setCurrentPage(p => Math.max(1, p - 1))}>Prev</button>
                                <span>Page {currentPage} of {totalPages}</span>
                                <button style={styles.btn} onClick={() => setCurrentPage(p => Math.min(totalPages, p + 1))}>Next</button>
                            </div>
                        </div>
                    </>
                ) : <p style={{ textAlign: 'center', color: '#666', padding: '40px' }}>Upload an Excel file to see the data table.</p>}
            </div>
        </div >
    );
};

export default ExcelToJsonConverter;