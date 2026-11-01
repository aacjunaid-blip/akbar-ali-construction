import React, { useState } from 'react';
import { Wallet, TrendingUp, TrendingDown, User, Building, LayoutDashboard, Landmark, FolderKanban, Receipt, Plus, Edit2, LogOut, FileSpreadsheet } from 'lucide-react';
import * as XLSX from 'xlsx';

function App() {
  const [currentUser, setCurrentUser] = useState(null);
  const [authMode, setAuthMode] = useState('login');
  
  const [registeredUsers, setRegisteredUsers] = useState([
    {
      name: "Akbar Ali",
      mobile: "03001234567",
      role: "Company Admin",
      password: "123",
      dp: ""
    }
  ]);

  const [authName, setAuthName] = useState('');
  const [authMobile, setAuthMobile] = useState('');
  const [authRole, setAuthRole] = useState('Project Engineer');
  const [authPassword, setAuthPassword] = useState('');
  const [authDp, setAuthDp] = useState('');
  const [authError, setAuthError] = useState('');

  const [activeTab, setActiveTab] = useState('dashboard');
  const [showEditProfile, setShowEditProfile] = useState(false);

  const [banks, setBanks] = useState([
    { id: 1, bankName: "Meezan Bank", accountNumber: "1234-5678-90", openingAmount: 500000, currentBalance: 500000 }
  ]);
  const [newBankName, setNewBankName] = useState('');
  const [newAccNo, setNewAccNo] = useState('');
  const [newOpening, setNewOpening] = useState('');

  const [projects, setProjects] = useState([
    { id: 1, name: "Guard Room HIT", location: "HIT" }
  ]);
  const [newProjName, setNewProjName] = useState('');
  const [selectedProject, setSelectedProject] = useState(null);

  const [transactions, setTransactions] = useState([]);
  const [showAddTxModal, setShowAddTxModal] = useState(false);
  const [txType, setTxType] = useState('Expense');
  const [txAmount, setTxAmount] = useState('');
  const [txBank, setTxBank] = useState('');
  const [txProject, setTxProject] = useState('');
  const [txSource, setTxSource] = useState('');
  const [txDate, setTxDate] = useState('');

  const handleSignUp = (e) => {
    e.preventDefault();
    setAuthError('');

    const cleanMobile = authMobile.trim();
    const mobileRegex = /^03\d{9}$/;
    if (!mobileRegex.test(cleanMobile)) {
      setAuthError("❌ Ghalat format! Mobile number 11 digits ka aur '03' se shuru hona chahiye (jaise: 03001234567).");
      return;
    }

    const exists = registeredUsers.some(u => u.mobile === cleanMobile);
    if (exists) {
      setAuthError("❌ Yeh mobile number pehle se registered hai! Ek number se ek hi account ban sakta hai.");
      return;
    }

    const newUser = {
      name: authName,
      mobile: cleanMobile,
      role: authRole,
      password: authPassword,
      dp: authDp
    };

    setRegisteredUsers([...registeredUsers, newUser]);
    setCurrentUser(newUser);
    setAuthName('');
    setAuthMobile('');
    setAuthPassword('');
    setAuthDp('');
  };

  const handleLogin = (e) => {
    e.preventDefault();
    setAuthError('');

    const cleanMobile = authMobile.trim();
    const user = registeredUsers.find(u => u.mobile === cleanMobile && u.password === authPassword);
    
    if (user) {
      setCurrentUser(user);
      setAuthMobile('');
      setAuthPassword('');
    } else {
      setAuthError("❌ Mobile Number ya Password galat hai!");
    }
  };

  const handleImageUpload = (e, isAuth = false) => {
    const file = e.target.files[0];
    if (file) {
      const reader = new FileReader();
      reader.onloadend = () => {
        if (isAuth) {
          setAuthDp(reader.result);
        } else {
          const updatedUser = { ...currentUser, dp: reader.result };
          setCurrentUser(updatedUser);
          setRegisteredUsers(registeredUsers.map(u => u.mobile === currentUser.mobile ? updatedUser : u));
        }
      };
      reader.readAsDataURL(file);
    }
  };

  const handleAddBank = (e) => {
    e.preventDefault();
    if (!newBankName || !newOpening) return;
    const opening = parseFloat(newOpening);
    const newB = {
      id: Date.now(),
      bankName: newBankName,
      accountNumber: newAccNo,
      openingAmount: opening,
      currentBalance: opening
    };
    setBanks([...banks, newB]);
    setNewBankName('');
    setNewAccNo('');
    setNewOpening('');
  };

  const handleAddProject = (e) => {
    e.preventDefault();
    if (!newProjName) return;
    const newP = { id: Date.now(), name: newProjName };
    setProjects([...projects, newP]);
    setNewProjName('');
  };

  const handleAddTransaction = (e) => {
    e.preventDefault();
    if (!txAmount || !txBank) return;
    const amount = parseFloat(txAmount);

    const updatedBanks = banks.map(b => {
      if (b.bankName === txBank) {
        let newBal = b.currentBalance;
        if (txType === 'Income') newBal += amount;
        else if (txType === 'Expense') newBal -= amount;
        return { ...b, currentBalance: newBal };
      }
      return b;
    });
    setBanks(updatedBanks);

    const newTx = {
      id: Date.now(),
      type: txType,
      amount: amount,
      bankName: txBank,
      projectName: txProject || 'General',
      sourceOrCategory: txSource,
      date: txDate || new Date().toISOString().split('T')[0]
    };

    setTransactions([newTx, ...transactions]);
    setShowAddTxModal(false);
    setTxAmount('');
    setTxSource('');
  };

  // Direct Excel Export Function for Marked Table Area
  const exportToExcel = (projectName, projectTransactions) => {
    const dataToExport = projectTransactions.map(t => ({
      'Date': t.date,
      'Notes / Description': t.sourceOrCategory,
      'Bank': t.bankName,
      'Type': t.type,
      'Expenses (Rs.)': t.amount
    }));

    const worksheet = XLSX.utils.json_to_sheet(dataToExport);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, projectName);

    XLSX.writeFile(workbook, `${projectName}_Ledger.xlsx`);
  };

  let totalIncome = transactions.filter(t => t.type === 'Income').reduce((acc, t) => acc + t.amount, 0);
  let totalExpense = transactions.filter(t => t.type === 'Expense').reduce((acc, t) => acc + t.amount, 0);
  let totalBalance = banks.reduce((acc, b) => acc + b.currentBalance, 0);

  if (!currentUser) {
    return (
      <div style={{ display: 'flex', justifyContent: 'center', alignItems: 'center', minHeight: '100vh', backgroundColor: '#f0fdf4', fontFamily: 'Arial, sans-serif' }}>
        <div style={{ backgroundColor: '#fff', padding: '35px', borderRadius: '12px', boxShadow: '0 4px 20px rgba(0,0,0,0.1)', width: '400px', borderTop: '6px solid #16a34a' }}>
          
          <div style={{ textAlign: 'center', marginBottom: '25px' }}>
            <h2 style={{ color: '#14532d', margin: '0 0 5px 0' }}>Akbar Ali & Co.</h2>
            <p style={{ color: '#666', fontSize: '14px', margin: 0 }}>Construction & Finance Portal</p>
          </div>

          {authError && (
            <div style={{ backgroundColor: '#fef2f2', color: '#dc2626', padding: '10px', borderRadius: '6px', fontSize: '13px', marginBottom: '15px', border: '1px solid #fecaca' }}>
              {authError}
            </div>
          )}

          {authMode === 'signup' ? (
            <form onSubmit={handleSignUp}>
              <h3 style={{ color: '#166534', marginTop: 0 }}>Create New Account</h3>
              
              <div style={{ marginBottom: '12px' }}>
                <label style={{ fontSize: '12px', fontWeight: 'bold', color: '#333' }}>Full Name</label>
                <input type="text" placeholder="e.g. Ali Raza" value={authName} onChange={(e) => setAuthName(e.target.value)} style={{ width: '100%', padding: '10px', marginTop: '4px', borderRadius: '6px', border: '1px solid #ccc', boxSizing: 'border-box' }} required />
              </div>

              <div style={{ marginBottom: '12px' }}>
                <label style={{ fontSize: '12px', fontWeight: 'bold', color: '#333' }}>Mobile Number (11 Digits)</label>
                <input type="text" maxLength="11" placeholder="03001234567" value={authMobile} onChange={(e) => setAuthMobile(e.target.value)} style={{ width: '100%', padding: '10px', marginTop: '4px', borderRadius: '6px', border: '1px solid #ccc', boxSizing: 'border-box' }} required />
              </div>

              <div style={{ marginBottom: '12px' }}>
                <label style={{ fontSize: '12px', fontWeight: 'bold', color: '#333' }}>Designation / Role</label>
                <input type="text" placeholder="e.g. Site Engineer" value={authRole} onChange={(e) => setAuthRole(e.target.value)} style={{ width: '100%', padding: '10px', marginTop: '4px', borderRadius: '6px', border: '1px solid #ccc', boxSizing: 'border-box' }} required />
              </div>

              <div style={{ marginBottom: '12px' }}>
                <label style={{ fontSize: '12px', fontWeight: 'bold', color: '#333' }}>Password</label>
                <input type="password" placeholder="••••••••" value={authPassword} onChange={(e) => setAuthPassword(e.target.value)} style={{ width: '100%', padding: '10px', marginTop: '4px', borderRadius: '6px', border: '1px solid #ccc', boxSizing: 'border-box' }} required />
              </div>

              <div style={{ marginBottom: '15px' }}>
                <label style={{ fontSize: '12px', fontWeight: 'bold', color: '#333' }}>Upload Profile Picture (Optional)</label>
                <input type="file" accept="image/*" onChange={(e) => handleImageUpload(e, true)} style={{ width: '100%', marginTop: '4px' }} />
              </div>

              <button type="submit" style={{ width: '100%', backgroundColor: '#16a34a', color: '#fff', border: 'none', padding: '12px', borderRadius: '6px', fontWeight: 'bold', cursor: 'pointer', fontSize: '15px' }}>
                Create Account
              </button>

              <div style={{ textAlign: 'center', marginTop: '15px' }}>
                <span style={{ fontSize: '13px', color: '#666' }}>Pehle se account hai? </span>
                <button type="button" onClick={() => { setAuthMode('login'); setAuthError(''); }} style={{ background: 'none', border: 'none', color: '#166534', fontWeight: 'bold', cursor: 'pointer' }}>
                  Login Karein
                </button>
              </div>
            </form>
          ) : (
            <form onSubmit={handleLogin}>
              <h3 style={{ color: '#166534', marginTop: 0 }}>Login</h3>
              
              <div style={{ marginBottom: '15px' }}>
                <label style={{ fontSize: '12px', fontWeight: 'bold', color: '#333' }}>Mobile Number (11 Digits)</label>
                <input type="text" maxLength="11" placeholder="03001234567" value={authMobile} onChange={(e) => setAuthMobile(e.target.value)} style={{ width: '100%', padding: '10px', marginTop: '4px', borderRadius: '6px', border: '1px solid #ccc', boxSizing: 'border-box' }} required />
              </div>

              <div style={{ marginBottom: '20px' }}>
                <label style={{ fontSize: '12px', fontWeight: 'bold', color: '#333' }}>Password</label>
                <input type="password" placeholder="••••••••" value={authPassword} onChange={(e) => setAuthPassword(e.target.value)} style={{ width: '100%', padding: '10px', marginTop: '4px', borderRadius: '6px', border: '1px solid #ccc', boxSizing: 'border-box' }} required />
              </div>

              <button type="submit" style={{ width: '100%', backgroundColor: '#166534', color: '#fff', border: 'none', padding: '12px', borderRadius: '6px', fontWeight: 'bold', cursor: 'pointer', fontSize: '15px' }}>
                Log In
              </button>

              <div style={{ textAlign: 'center', marginTop: '20px', borderTop: '1px solid #eee', paddingTop: '15px' }}>
                <button type="button" onClick={() => { setAuthMode('signup'); setAuthError(''); }} style={{ backgroundColor: '#22c55e', color: '#fff', border: 'none', padding: '10px 20px', borderRadius: '6px', fontWeight: 'bold', cursor: 'pointer' }}>
                  Create New Account
                </button>
              </div>
            </form>
          )}

        </div>
      </div>
    );
  }

  return (
    <div style={{ display: 'flex', height: '100vh', backgroundColor: '#f0fdf4', fontFamily: 'Arial, sans-serif' }}>
      
      {/* Sidebar */}
      <div style={{ width: '260px', backgroundColor: '#14532d', color: '#fff', display: 'flex', flexDirection: 'column', padding: '20px', overflowY: 'auto' }}>
        
        <div style={{ display: 'flex', alignItems: 'center', gap: '12px', marginBottom: '25px', paddingBottom: '15px', borderBottom: '1px solid #166534' }}>
          <div onClick={() => setShowEditProfile(true)} style={{ position: 'relative', cursor: 'pointer' }}>
            {currentUser.dp ? (
              <img src={currentUser.dp} alt="DP" style={{ width: '50px', height: '50px', borderRadius: '50%', objectFit: 'cover', border: '2px solid #86efac' }} />
            ) : (
              <div style={{ width: '50px', height: '50px', borderRadius: '50%', backgroundColor: '#166534', display: 'flex', alignItems: 'center', justifyContent: 'center', border: '2px solid #86efac', color: '#fff' }}>
                <User size={26} />
              </div>
            )}
            <div style={{ position: 'absolute', bottom: 0, right: 0, backgroundColor: '#22c55e', borderRadius: '50%', padding: '2px' }}>
              <Edit2 size={12} color="#fff" />
            </div>
          </div>
          <div style={{ flex: 1 }}>
            <h4 style={{ margin: '0', fontSize: '15px', color: '#86efac' }}>{currentUser.name}</h4>
            <p style={{ margin: '2px 0 0 0', fontSize: '12px', color: '#cbd5e1' }}>{currentUser.role}</p>
          </div>
        </div>

        <h2 style={{ fontSize: '18px', marginBottom: '25px', color: '#86efac' }}>Akbar Ali & Co.</h2>
        
        <div style={{ display: 'flex', flexDirection: 'column', gap: '10px', flex: 1 }}>
          <div onClick={() => { setActiveTab('dashboard'); setSelectedProject(null); }} style={{ display: 'flex', alignItems: 'center', gap: '10px', cursor: 'pointer', padding: '10px', borderRadius: '8px', backgroundColor: activeTab === 'dashboard' ? '#166534' : 'transparent' }}>
            <LayoutDashboard size={18} /> Dashboard
          </div>
          <div onClick={() => setActiveTab('banks')} style={{ display: 'flex', alignItems: 'center', gap: '10px', cursor: 'pointer', padding: '10px', borderRadius: '8px', backgroundColor: activeTab === 'banks' ? '#166534' : 'transparent' }}>
            <Landmark size={18} /> Bank Accounts
          </div>
          <div onClick={() => { setActiveTab('projects'); setSelectedProject(null); }} style={{ display: 'flex', alignItems: 'center', gap: '10px', cursor: 'pointer', padding: '10px', borderRadius: '8px', backgroundColor: activeTab === 'projects' ? '#166534' : 'transparent' }}>
            <FolderKanban size={18} /> Projects & Ledger
          </div>
          <div onClick={() => setActiveTab('transactions')} style={{ display: 'flex', alignItems: 'center', gap: '10px', cursor: 'pointer', padding: '10px', borderRadius: '8px', backgroundColor: activeTab === 'transactions' ? '#166534' : 'transparent' }}>
            <Receipt size={18} /> All Transactions
          </div>
        </div>

        <button onClick={() => setCurrentUser(null)} style={{ display: 'flex', alignItems: 'center', gap: '10px', backgroundColor: '#dc2626', color: '#fff', border: 'none', padding: '10px', borderRadius: '8px', cursor: 'pointer', marginTop: 'auto' }}>
          <LogOut size={18} /> Log Out
        </button>
      </div>

      {/* Main Content Area */}
      <div style={{ flex: 1, display: 'flex', flexDirection: 'column', overflowY: 'auto' }}>
        
        <header style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', backgroundColor: '#ffffff', padding: '15px 30px', borderBottom: '1px solid #dcfce7' }}>
          <h3 style={{ margin: 0, color: '#166534' }}>Construction & Finance Management Software</h3>
          <button onClick={() => setShowAddTxModal(true)} style={{ backgroundColor: '#16a34a', color: '#fff', border: 'none', padding: '10px 20px', borderRadius: '6px', cursor: 'pointer', fontWeight: 'bold', display: 'flex', alignItems: 'center', gap: '8px' }}>
            <Plus size={18} /> Add New Entry
          </button>
        </header>

        <div style={{ padding: '30px' }}>

          {activeTab === 'dashboard' && (
            <div>
              <h2 style={{ color: '#14532d', marginBottom: '5px' }}>Dashboard</h2>
              <p style={{ color: '#666', marginBottom: '25px' }}>Overview of all accounts and financial position</p>

              <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3, 1fr)', gap: '20px', marginBottom: '30px' }}>
                <div style={{ backgroundColor: '#dcfce7', borderLeft: '6px solid #16a34a', padding: '20px', borderRadius: '8px' }}>
                  <h4 style={{ margin: '0 0 10px 0', color: '#14532d' }}>REMAINING BALANCE</h4>
                  <h2 style={{ margin: 0, color: '#166534' }}>Rs. {totalBalance.toLocaleString()}</h2>
                </div>
                <div style={{ backgroundColor: '#fef2f2', borderLeft: '6px solid #dc2626', padding: '20px', borderRadius: '8px' }}>
                  <h4 style={{ margin: '0 0 10px 0', color: '#991b1b' }}>TOTAL EXPENSE</h4>
                  <h2 style={{ margin: 0, color: '#dc2626' }}>Rs. {totalExpense.toLocaleString()}</h2>
                </div>
                <div style={{ backgroundColor: '#f0fdf4', borderLeft: '6px solid #2563eb', padding: '20px', borderRadius: '8px' }}>
                  <h4 style={{ margin: '0 0 10px 0', color: '#1e40af' }}>TOTAL INCOME</h4>
                  <h2 style={{ margin: 0, color: '#2563eb' }}>Rs. {totalIncome.toLocaleString()}</h2>
                </div>
              </div>
            </div>
          )}

          {activeTab === 'banks' && (
            <div>
              <h2 style={{ color: '#14532d', marginBottom: '15px' }}>Bank Accounts Management</h2>
              
              <form onSubmit={handleAddBank} style={{ display: 'flex', gap: '10px', marginBottom: '25px', backgroundColor: '#fff', padding: '15px', borderRadius: '8px', border: '1px solid #dcfce7' }}>
                <input type="text" placeholder="Bank Name" value={newBankName} onChange={(e) => setNewBankName(e.target.value)} style={{ padding: '8px', flex: 1, borderRadius: '4px', border: '1px solid #ccc' }} required />
                <input type="text" placeholder="Account Number" value={newAccNo} onChange={(e) => setNewAccNo(e.target.value)} style={{ padding: '8px', flex: 1, borderRadius: '4px', border: '1px solid #ccc' }} />
                <input type="number" placeholder="Opening Amount" value={newOpening} onChange={(e) => setNewOpening(e.target.value)} style={{ padding: '8px', flex: 1, borderRadius: '4px', border: '1px solid #ccc' }} required />
                <button type="submit" style={{ backgroundColor: '#16a34a', color: '#fff', border: 'none', padding: '8px 15px', borderRadius: '4px', cursor: 'pointer' }}>Add Bank</button>
              </form>

              <table style={{ width: '100%', backgroundColor: '#fff', borderCollapse: 'collapse', borderRadius: '8px', overflow: 'hidden' }}>
                <thead>
                  <tr style={{ backgroundColor: '#dcfce7', color: '#14532d', textAlign: 'left' }}>
                    <th style={{ padding: '12px' }}>Bank Name</th>
                    <th style={{ padding: '12px' }}>Account Number</th>
                    <th style={{ padding: '12px' }}>Opening Amount</th>
                    <th style={{ padding: '12px' }}>Current Balance</th>
                  </tr>
                </thead>
                <tbody>
                  {banks.map(b => (
                    <tr key={b.id} style={{ borderBottom: '1px solid #eee' }}>
                      <td style={{ padding: '12px' }}>{b.bankName}</td>
                      <td style={{ padding: '12px' }}>{b.accountNumber || 'N/A'}</td>
                      <td style={{ padding: '12px' }}>Rs. {b.openingAmount.toLocaleString()}</td>
                      <td style={{ padding: '12px', fontWeight: 'bold', color: '#16a34a' }}>Rs. {b.currentBalance.toLocaleString()}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}

          {activeTab === 'projects' && !selectedProject && (
            <div>
              <h2 style={{ color: '#14532d', marginBottom: '15px' }}>Projects List</h2>
              
              <form onSubmit={handleAddProject} style={{ display: 'flex', gap: '10px', marginBottom: '25px', backgroundColor: '#fff', padding: '15px', borderRadius: '8px', border: '1px solid #dcfce7' }}>
                <input type="text" placeholder="Project Name (e.g. Guard Room HIT)" value={newProjName} onChange={(e) => setNewProjName(e.target.value)} style={{ padding: '8px', flex: 1, borderRadius: '4px', border: '1px solid #ccc' }} required />
                <button type="submit" style={{ backgroundColor: '#16a34a', color: '#fff', border: 'none', padding: '8px 15px', borderRadius: '4px', cursor: 'pointer' }}>Add Project</button>
              </form>

              <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3, 1fr)', gap: '15px' }}>
                {projects.map(p => (
                  <div key={p.id} onClick={() => setSelectedProject(p)} style={{ backgroundColor: '#fff', padding: '20px', borderRadius: '8px', border: '1px solid #dcfce7', cursor: 'pointer', boxShadow: '0 2px 5px rgba(0,0,0,0.05)' }}>
                    <h3 style={{ margin: '0 0 10px 0', color: '#14532d' }}>{p.name}</h3>
                    <p style={{ margin: 0, fontSize: '13px', color: '#666' }}>Click to open project ledger</p>
                  </div>
                ))}
              </div>
            </div>
          )}

          {selectedProject && (
            <div>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '15px' }}>
                <button onClick={() => setSelectedProject(null)} style={{ backgroundColor: '#64748b', color: '#fff', border: 'none', padding: '6px 12px', borderRadius: '4px', cursor: 'pointer' }}>← Back to Projects</button>
                
                {/* Excel Export Button replacing Print */}
                <button onClick={() => {
                  const projTx = transactions.filter(t => t.projectName === selectedProject.name);
                  exportToExcel(selectedProject.name, projTx);
                }} style={{ backgroundColor: '#16a34a', color: '#fff', border: 'none', padding: '10px 20px', borderRadius: '6px', cursor: 'pointer', fontWeight: 'bold', display: 'flex', alignItems: 'center', gap: '8px' }}>
                  <FileSpreadsheet size={18} /> Export to Excel ({selectedProject.name})
                </button>
              </div>

              {/* Marked Table Area */}
              <div style={{ backgroundColor: '#fff', padding: '20px', borderRadius: '8px', border: '1px solid #dcfce7' }}>
                <h2 style={{ textAlign: 'center', color: '#14532d', marginBottom: '20px', borderBottom: '2px solid #16a34a', paddingBottom: '10px' }}>{selectedProject.name}</h2>
                
                <table style={{ width: '100%', backgroundColor: '#fff', borderCollapse: 'collapse', borderRadius: '8px', overflow: 'hidden' }}>
                  <thead>
                    <tr style={{ backgroundColor: '#dcfce7', color: '#14532d', textAlign: 'left' }}>
                      <th style={{ padding: '12px', border: '1px solid #cbd5e1' }}>Date</th>
                      <th style={{ padding: '12px', border: '1px solid #cbd5e1' }}>Notes / Description</th>
                      <th style={{ padding: '12px', border: '1px solid #cbd5e1' }}>Bank</th>
                      <th style={{ padding: '12px', border: '1px solid #cbd5e1' }}>Type</th>
                      <th style={{ padding: '12px', border: '1px solid #cbd5e1' }}>Expenses (Rs.)</th>
                    </tr>
                  </thead>
                  <tbody>
                    {transactions.filter(t => t.projectName === selectedProject.name).length === 0 ? (
                      <tr><td colSpan="5" style={{ padding: '20px', textAlign: 'center', color: '#666', border: '1px solid #cbd5e1' }}>Koi transaction nahi hai is project mein.</td></tr>
                    ) : (
                      transactions.filter(t => t.projectName === selectedProject.name).map(t => (
                        <tr key={t.id} style={{ borderBottom: '1px solid #eee' }}>
                          <td style={{ padding: '12px', border: '1px solid #cbd5e1' }}>{t.date}</td>
                          <td style={{ padding: '12px', border: '1px solid #cbd5e1' }}>{t.sourceOrCategory}</td>
                          <td style={{ padding: '12px', border: '1px solid #cbd5e1' }}>{t.bankName}</td>
                          <td style={{ padding: '12px', border: '1px solid #cbd5e1', color: t.type === 'Income' ? '#16a34a' : '#dc2626', fontWeight: 'bold' }}>{t.type}</td>
                          <td style={{ padding: '12px', border: '1px solid #cbd5e1', fontWeight: 'bold' }}>Rs. {t.amount.toLocaleString()}</td>
                        </tr>
                      ))
                    )}
                  </tbody>
                </table>
              </div>
            </div>
          )}

          {activeTab === 'transactions' && (
            <div>
              <h2 style={{ color: '#14532d', marginBottom: '15px' }}>All Transactions Ledger</h2>
              <table style={{ width: '100%', backgroundColor: '#fff', borderCollapse: 'collapse', borderRadius: '8px', overflow: 'hidden' }}>
                <thead>
                  <tr style={{ backgroundColor: '#dcfce7', color: '#14532d', textAlign: 'left' }}>
                    <th style={{ padding: '12px' }}>Date</th>
                    <th style={{ padding: '12px' }}>Type</th>
                    <th style={{ padding: '12px' }}>Project</th>
                    <th style={{ padding: '12px' }}>Source/Category</th>
                    <th style={{ padding: '12px' }}>Bank</th>
                    <th style={{ padding: '12px' }}>Amount</th>
                  </tr>
                </thead>
                <tbody>
                  {transactions.length === 0 ? (
                    <tr><td colSpan="6" style={{ padding: '20px', textAlign: 'center', color: '#666' }}>Abhi tak koi entry nahi ki gayi.</td></tr>
                  ) : (
                    transactions.map(t => (
                      <tr key={t.id} style={{ borderBottom: '1px solid #eee' }}>
                        <td style={{ padding: '12px' }}>{t.date}</td>
                        <td style={{ padding: '12px', color: t.type === 'Income' ? '#16a34a' : '#dc2626', fontWeight: 'bold' }}>{t.type}</td>
                        <td style={{ padding: '12px' }}>{t.projectName}</td>
                        <td style={{ padding: '12px' }}>{t.sourceOrCategory}</td>
                        <td style={{ padding: '12px' }}>{t.bankName}</td>
                        <td style={{ padding: '12px', fontWeight: 'bold' }}>Rs. {t.amount.toLocaleString()}</td>
                      </tr>
                    ))
                  )}
                </tbody>
              </table>
            </div>
          )}

        </div>

      </div>

      {showEditProfile && (
        <div style={{ position: 'fixed', top: 0, left: 0, width: '100%', height: '100%', backgroundColor: 'rgba(0,0,0,0.5)', display: 'flex', justifyContent: 'center', alignItems: 'center', zIndex: 100 }}>
          <div style={{ backgroundColor: '#fff', padding: '25px', borderRadius: '8px', width: '350px' }}>
            <h3 style={{ marginTop: 0, color: '#14532d' }}>Edit Employee Profile</h3>
            <div style={{ marginBottom: '15px' }}>
              <label style={{ fontSize: '13px', display: 'block', marginBottom: '5px' }}>Name</label>
              <input type="text" value={currentUser.name} onChange={(e) => setCurrentUser({...currentUser, name: e.target.value})} style={{ width: '100%', padding: '8px', borderRadius: '4px', border: '1px solid #ccc' }} />
            </div>
            <div style={{ marginBottom: '15px' }}>
              <label style={{ fontSize: '13px', display: 'block', marginBottom: '5px' }}>Designation / Role</label>
              <input type="text" value={currentUser.role} onChange={(e) => setCurrentUser({...currentUser, role: e.target.value})} style={{ width: '100%', padding: '8px', borderRadius: '4px', border: '1px solid #ccc' }} />
            </div>
            <div style={{ marginBottom: '15px' }}>
              <label style={{ fontSize: '13px', display: 'block', marginBottom: '5px' }}>Contact Number</label>
              <input type="text" value={currentUser.mobile} disabled style={{ width: '100%', padding: '8px', borderRadius: '4px', border: '1px solid #ccc', backgroundColor: '#eee' }} />
            </div>
            <div style={{ marginBottom: '15px' }}>
              <label style={{ fontSize: '13px', display: 'block', marginBottom: '5px' }}>Upload New DP</label>
              <input type="file" accept="image/*" onChange={(e) => handleImageUpload(e, false)} style={{ width: '100%' }} />
            </div>
            <div style={{ display: 'flex', justifyContent: 'flex-end', gap: '10px' }}>
              <button onClick={() => setShowEditProfile(false)} style={{ backgroundColor: '#16a34a', color: '#fff', border: 'none', padding: '8px 15px', borderRadius: '4px', cursor: 'pointer' }}>Done</button>
            </div>
          </div>
        </div>
      )}

      {showAddTxModal && (
        <div style={{ position: 'fixed', top: 0, left: 0, width: '100%', height: '100%', backgroundColor: 'rgba(0,0,0,0.5)', display: 'flex', justifyContent: 'center', alignItems: 'center', zIndex: 100 }}>
          <form onSubmit={handleAddTransaction} style={{ backgroundColor: '#fff', padding: '25px', borderRadius: '8px', width: '400px' }}>
            <h3 style={{ marginTop: 0, color: '#14532d' }}>Add New Financial Entry</h3>
            
            <div style={{ marginBottom: '12px' }}>
              <label style={{ fontSize: '13px', display: 'block', marginBottom: '5px' }}>Type</label>
              <select value={txType} onChange={(e) => setTxType(e.target.value)} style={{ width: '100%', padding: '8px', borderRadius: '4px', border: '1px solid #ccc' }}>
                <option value="Expense">Expense</option>
                <option value="Income">Income</option>
              </select>
            </div>

            <div style={{ marginBottom: '12px' }}>
              <label style={{ fontSize: '13px', display: 'block', marginBottom: '5px' }}>Amount (Rs.)</label>
              <input type="number" value={txAmount} onChange={(e) => setTxAmount(e.target.value)} style={{ width: '100%', padding: '8px', borderRadius: '4px', border: '1px solid #ccc' }} required />
            </div>

            <div style={{ marginBottom: '12px' }}>
              <label style={{ fontSize: '13px', display: 'block', marginBottom: '5px' }}>Date</label>
              <input type="date" value={txDate} onChange={(e) => setTxDate(e.target.value)} style={{ width: '100%', padding: '8px', borderRadius: '4px', border: '1px solid #ccc' }} />
            </div>

            <div style={{ marginBottom: '12px' }}>
              <label style={{ flexibility: '1', fontSize: '13px', display: 'block', marginBottom: '5px' }}>Bank Name (Relevant Bank)</label>
              <select value={txBank} onChange={(e) => setTxBank(e.target.value)} style={{ width: '100%', padding: '8px', borderRadius: '4px', border: '1px solid #ccc' }} required>
                <option value="">Select Bank</option>
                {banks.map(b => (
                  <option key={b.id} value={b.bankName}>{b.bankName} (Bal: Rs. {b.currentBalance})</option>
                ))}
              </select>
            </div>

            <div style={{ marginBottom: '12px' }}>
              <label style={{ fontSize: '13px', display: 'block', marginBottom: '5px' }}>Project Name</label>
              <select value={txProject} onChange={(e) => setTxProject(e.target.value)} style={{ width: '100%', padding: '8px', borderRadius: '4px', border: '1px solid #ccc' }} required>
                <option value="">Select Project</option>
                {projects.map(p => (
                  <option key={p.id} value={p.name}>{p.name}</option>
                ))}
              </select>
            </div>

            <div style={{ marginBottom: '15px' }}>
              <label style={{ fontSize: '13px', display: 'block', marginBottom: '5px' }}>Notes / Description</label>
              <input type="text" value={txSource} onChange={(e) => setTxSource(e.target.value)} placeholder="e.g. Amir Bricks from ASK A/C / Lunch" style={{ width: '100%', padding: '8px', borderRadius: '4px', border: '1px solid #ccc' }} required />
            </div>

            <div style={{ display: 'flex', justifyContent: 'flex-end', gap: '10px' }}>
              <button type="button" onClick={() => setShowAddTxModal(false)} style={{ backgroundColor: '#64748b', color: '#fff', border: 'none', padding: '8px 15px', borderRadius: '4px', cursor: 'pointer' }}>Cancel</button>
              <button type="submit" style={{ backgroundColor: '#16a34a', color: '#fff', border: 'none', padding: '8px 15px', borderRadius: '4px', cursor: 'pointer' }}>Save Entry</button>
            </div>
          </form>
        </div>
      )}

    </div>
  );
}

export default App;