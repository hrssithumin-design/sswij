import React, { useEffect, useRef, useState } from "react";
import { createRoot } from "react-dom/client";
import { Document, Packer, Paragraph, TextRun, HeadingLevel } from "docx";
import { supabase, supabaseConfigured } from "./supabase";
import "./style.css";

function AuthGate() {
  const [session, setSession] = useState(null);
  const [loading, setLoading] = useState(true);
  useEffect(() => {
    if (!supabaseConfigured) { setLoading(false); return; }
    supabase.auth.getSession().then(({data}) => { setSession(data.session); setLoading(false); });
    const {data:{subscription}} = supabase.auth.onAuthStateChange((_event, next) => setSession(next));
    return () => subscription.unsubscribe();
  }, []);
  if (loading) return <div className="auth-screen"><div className="auth-card">Loading WordCraft…</div></div>;
  if (!supabaseConfigured) return <SetupScreen />;
  if (!session) return <AuthScreen />;
  return <EditorApp user={session.user} />;
}

function Brand() {
  return <div className="brand"><div className="brand-icon">W</div><div><strong>WordCraft</strong><small>YOUR DOCUMENT WORKSPACE</small></div></div>;
}
function SetupScreen() {
  return <div className="auth-screen"><div className="auth-card"><Brand/><h1>Connect your workspace</h1><p>WordCraft login uses Supabase authentication. Add your project URL and public anon/publishable key to a <code>.env</code> file, then restart the app.</p><p className="hint">See <b>README.md</b> for the quick setup and database security SQL. Never put a service-role key in this app.</p></div></div>;
}
function AuthScreen() {
  const [mode,setMode] = useState("login");
  const [email,setEmail] = useState("");
  const [password,setPassword] = useState("");
  const [busy,setBusy] = useState(false);
  const [message,setMessage] = useState("");
  const [error,setError] = useState("");
  async function submit(e) {
    e.preventDefault(); setBusy(true); setError(""); setMessage("");
    try {
      if (mode === "signup") {
        const {data,error} = await supabase.auth.signUp({email,password,options:{emailRedirectTo:window.location.origin}});
        if (error) throw error;
        setMessage(data.session ? "Your account is ready." : "Check your email to confirm your account, then sign in.");
      } else if (mode === "login") {
        const {error} = await supabase.auth.signInWithPassword({email,password});
        if (error) throw error;
      } else {
        const {error} = await supabase.auth.resetPasswordForEmail(email,{redirectTo:window.location.origin});
        if (error) throw error;
        setMessage("Password reset link sent. Check your email.");
      }
    } catch(e) { setError(e.message || "Something went wrong."); }
    finally { setBusy(false); }
  }
  async function google() {
    setError(""); setBusy(true);
    const {error} = await supabase.auth.signInWithOAuth({provider:"google",options:{redirectTo:window.location.origin}});
    if (error) {setError(error.message);setBusy(false);}
  }
  const heading = mode==="login" ? "Welcome back" : mode==="signup" ? "Create your account" : "Reset your password";
  return <div className="auth-screen"><div className="auth-card">
    <Brand/><h1>{heading}</h1><p className="auth-sub">{mode==="login"?"Sign in to access your documents.":mode==="signup"?"Create an account to keep your work in one place.":"We’ll email you a secure reset link."}</p>
    <form onSubmit={submit}>
      <label>Email address</label><input type="email" value={email} onChange={e=>setEmail(e.target.value)} placeholder="you@example.com" required autoComplete="email"/>
      {mode!=="reset" && <><label>Password</label><input type="password" value={password} onChange={e=>setPassword(e.target.value)} placeholder="At least 6 characters" required minLength={6} autoComplete={mode==="login"?"current-password":"new-password"}/></>}
      <button className="auth-submit" disabled={busy}>{busy?"Please wait…":mode==="login"?"Sign in":mode==="signup"?"Create account":"Send reset link"}</button>
    </form>
    {mode!=="reset" && <><div className="or"><span>or</span></div><button className="google-btn" onClick={google} disabled={busy}><b>G</b> Continue with Google</button></>}
    {error && <p className="auth-error">{error}</p>}{message && <p className="auth-message">{message}</p>}
    <div className="auth-links">
      {mode==="login" ? <><button onClick={()=>setMode("signup")}>Create an account</button><button onClick={()=>setMode("reset")}>Forgot password?</button></> :
       <button onClick={()=>{setMode("login");setError("");setMessage("");}}>Back to sign in</button>}
    </div>
    <p className="secure-note">Your account is secured by Supabase Authentication.</p>
  </div></div>;
}

const initial = `<h1>Welcome to WordCraft</h1><p>Your ideas deserve a beautiful page. Start writing here.</p><p><br></p><p><b>WordCraft</b> is your simple, focused document workspace.</p>`;
function EditorApp({user}) {
  const editor = useRef(null), fileInput = useRef(null), imageInput = useRef(null);
  const [name,setName] = useState("Untitled document");
  const [saved,setSaved] = useState("All changes saved");
  const [font,setFont] = useState("Arial"), [size,setSize] = useState("12");
  const [wordCount,setWordCount] = useState(0), [zoom,setZoom] = useState(100);
  const [showFind,setShowFind] = useState(false), [findText,setFindText] = useState(""), [replaceText,setReplaceText] = useState("");
  const [documents,setDocuments] = useState([]);
  const [showDocs,setShowDocs] = useState(false);
  const [docId,setDocId] = useState(null);
  const localKey = `wordcraft:${user.id}:draft`;
  const titleKey = `wordcraft:${user.id}:title`;

  useEffect(() => {
    let active = true;
    (async()=>{
      const {data,error} = await supabase.from("documents").select("id,title,updated_at").order("updated_at",{ascending:false});
      if (active && !error) setDocuments(data || []);
      const {data:local} = await supabase.from("documents").select("id,title,content").order("updated_at",{ascending:false}).limit(1);
      if (!active) return;
      if (local?.length) {
        setDocId(local[0].id); setName(local[0].title);
        if(editor.current) editor.current.innerHTML=local[0].content || initial;
      } else {
        const draft=localStorage.getItem(localKey);
        if(editor.current) editor.current.innerHTML=draft || initial;
        setName(localStorage.getItem(titleKey) || "Untitled document");
      }
      updateCount();
    })();
    return ()=>{active=false;};
  },[]);
  function updateCount() {
    if(!editor.current)return;
    setWordCount(editor.current.innerText.trim().split(/\s+/).filter(Boolean).length);
  }
  async function save() {
    if(!editor.current)return;
    const content=editor.current.innerHTML;
    localStorage.setItem(localKey,content); localStorage.setItem(titleKey,name);
    setSaved("Saving…");
    const payload={user_id:user.id,title:name,content,updated_at:new Date().toISOString()};
    let result;
    if(docId) result=await supabase.from("documents").update(payload).eq("id",docId).select("id").single();
    else result=await supabase.from("documents").insert(payload).select("id").single();
    if(result.error) {setSaved("Saved on this device · cloud sync unavailable");return;}
    if(result.data?.id) setDocId(result.data.id);
    setSaved("Saved to your account");
    const {data}=await supabase.from("documents").select("id,title,updated_at").order("updated_at",{ascending:false});
    setDocuments(data||[]);
  }
  useEffect(()=>{const timer=setTimeout(()=>save(),1000);return()=>clearTimeout(timer);},[name,wordCount]);
  function command(cmd,value=null){editor.current?.focus();document.execCommand(cmd,false,value);updateCount();}
  function insertTable(){
    const rows=Number(prompt("Number of rows?","3"))||3, cols=Number(prompt("Number of columns?","3"))||3;
    if(rows>15||cols>10||rows<1||cols<1)return alert("Choose 1–15 rows and 1–10 columns.");
    let html="<table><tbody>";for(let r=0;r<rows;r++){html+="<tr>";for(let c=0;c<cols;c++)html+="<td><br></td>";html+="</tr>";}html+="</tbody></table><p><br></p>";command("insertHTML",html);
  }
  function insertImage(file){if(!file)return;const reader=new FileReader();reader.onload=e=>{command("insertImage",e.target.result);};reader.readAsDataURL(file);}
  function download(blob,filename){const url=URL.createObjectURL(blob),a=document.createElement("a");a.href=url;a.download=filename;a.click();URL.revokeObjectURL(url);}
  function safeName(){return(name||"document").replace(/[\\/:*?"<>|]+/g,"-");}
  function escapeHtml(s){return s.replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));}
  function exportHTML(){const html=`<!doctype html><html><head><meta charset="utf-8"><title>${escapeHtml(name)}</title><style>body{font:12pt Arial;max-width:800px;margin:50px auto;line-height:1.6}table{border-collapse:collapse;width:100%}td{border:1px solid #777;padding:8px}img{max-width:100%}</style></head><body>${editor.current.innerHTML}</body></html>`;download(new Blob([html],{type:"text/html"}),`${safeName()}.html`);}
  async function exportDocx(){
    const children=[...editor.current.childNodes].map(n=>new Paragraph({children:[new TextRun(n.textContent||"")]}));
    const doc=new Document({sections:[{children:children.length?children:[new Paragraph("")]}]});
    download(await Packer.toBlob(doc),`${safeName()}.docx`);
  }
  function openFile(file){if(!file)return;const reader=new FileReader();reader.onload=e=>{editor.current.innerHTML=e.target.result;setName(file.name.replace(/\.[^.]+$/,""));setDocId(null);updateCount();};reader.readAsText(file);}
  function printDoc(){const w=window.open("","_blank");w.document.write(`<html><head><title>${escapeHtml(name)}</title><style>body{font:12pt Arial;line-height:1.6;margin:25mm}table{border-collapse:collapse;width:100%}td{border:1px solid #777;padding:8px}img{max-width:100%}</style></head><body>${editor.current.innerHTML}</body></html>`);w.document.close();w.focus();w.print();}
  function replaceAll(){if(!findText)return;editor.current.innerHTML=editor.current.innerHTML.split(findText).join(replaceText);updateCount();}
  async function loadDocument(id){const {data,error}=await supabase.from("documents").select("id,title,content").eq("id",id).single();if(error)return alert(error.message);setDocId(data.id);setName(data.title);editor.current.innerHTML=data.content||"";updateCount();setShowDocs(false);setSaved("Document loaded");}
  async function newDocument(){setDocId(null);setName("Untitled document");editor.current.innerHTML="<p><br></p>";updateCount();setShowDocs(false);}
  async function signOut(){await supabase.auth.signOut();}
  return <div className="app">
    <header className="topbar"><Brand/><input className="docname" value={name} onChange={e=>setName(e.target.value)} aria-label="Document name"/><div className="save-state"><span className="dot"></span>{saved}</div><button className="primary" onClick={save}>Save</button><button className="account-btn" onClick={signOut} title={user.email}>Sign out</button></header>
    <nav className="menu"><button onClick={()=>setShowDocs(v=>!v)}>My documents</button><button onClick={newDocument}>New document</button><button onClick={()=>fileInput.current.click()}>Open file</button><button onClick={exportHTML}>Export HTML</button><button onClick={exportDocx}>Download DOCX</button><button onClick={printDoc}>Print / PDF</button><button onClick={()=>setShowFind(v=>!v)}>Find & Replace</button><input ref={fileInput} type="file" accept=".html,.htm,.txt" hidden onChange={e=>openFile(e.target.files[0])}/></nav>
    {showDocs&&<div className="docs-panel"><div className="docs-heading"><b>Your documents</b><button onClick={()=>setShowDocs(false)}>Close</button></div>{documents.length?documents.map(d=><button className="doc-row" key={d.id} onClick={()=>loadDocument(d.id)}><span>📄 {d.title}</span><small>{new Date(d.updated_at).toLocaleString()}</small></button>):<p>No saved cloud documents yet. Start writing and your work will be saved here.</p>}</div>}
    {showFind&&<div className="findbar"><input placeholder="Find text" value={findText} onChange={e=>setFindText(e.target.value)}/><input placeholder="Replace with" value={replaceText} onChange={e=>setReplaceText(e.target.value)}/><button onClick={replaceAll}>Replace all</button><button className="quiet" onClick={()=>setShowFind(false)}>Close</button></div>}
    <section className="toolbar"><div className="toolgroup"><button title="Undo" onClick={()=>command("undo")}>↶</button><button title="Redo" onClick={()=>command("redo")}>↷</button></div><i/><select value={font} onChange={e=>{setFont(e.target.value);command("fontName",e.target.value)}}><option>Arial</option><option>Georgia</option><option>Times New Roman</option><option>Verdana</option><option>Calibri</option><option>Courier New</option></select><select className="size-select" value={size} onChange={e=>{setSize(e.target.value);command("fontSize",({8:1,10:2,12:3,14:4,18:5,24:6,36:7})[e.target.value]||3)}}>{["8","10","12","14","18","24","36"].map(x=><option key={x}>{x}</option>)}</select><i/><div className="toolgroup"><button onClick={()=>command("bold")}><b>B</b></button><button onClick={()=>command("italic")}><em>I</em></button><button onClick={()=>command("underline")}><u>U</u></button><button onClick={()=>command("strikeThrough")}><s>S</s></button></div><i/><div className="toolgroup"><button onClick={()=>command("justifyLeft")}>☰</button><button onClick={()=>command("justifyCenter")}>≡</button><button onClick={()=>command("justifyRight")}>☷</button></div><i/><div className="toolgroup"><button onClick={()=>command("insertUnorderedList")}>• List</button><button onClick={()=>command("insertOrderedList")}>1. List</button><button onClick={()=>command("outdent")}>⇤</button><button onClick={()=>command("indent")}>⇥</button></div><i/><div className="toolgroup"><button onClick={insertTable}>▦ Table</button><button onClick={()=>imageInput.current.click()}>▧ Image</button><input ref={imageInput} type="file" accept="image/*" hidden onChange={e=>insertImage(e.target.files[0])}/><button onClick={()=>command("formatBlock","H1")}>H1</button><button onClick={()=>command("formatBlock","H2")}>H2</button><button onClick={()=>command("removeFormat")}>Clear</button></div><div className="zoom"><button onClick={()=>setZoom(z=>Math.max(70,z-10))}>−</button><span>{zoom}%</span><button onClick={()=>setZoom(z=>Math.min(130,z+10))}>＋</button></div></section>
    <main className="workspace"><div className="page-wrap"><div className="page" style={{zoom:zoom/100}}><div ref={editor} className="editor" contentEditable suppressContentEditableWarning spellCheck onInput={()=>{updateCount();setSaved("Unsaved changes");}} onKeyDown={e=>{if((e.ctrlKey||e.metaKey)&&e.key==="s"){e.preventDefault();save();}}}/></div></div></main>
    <footer className="statusbar"><span>{wordCount} words</span><span>{user.email}</span><span>Cloud sync · A4 page</span><span className="spacer"/><span>WordCraft</span></footer>
  </div>;
}
createRoot(document.getElementById("root")).render(<AuthGate/>);
