import { initializeApp } from "https://www.gstatic.com/firebasejs/10.8.0/firebase-app.js";
import { getAuth, signInWithEmailAndPassword, signOut, onAuthStateChanged } from "https://www.gstatic.com/firebasejs/10.8.0/firebase-auth.js";
import { getFirestore, collection, addDoc, getDocs, deleteDoc, doc, query, orderBy } from "https://www.gstatic.com/firebasejs/10.8.0/firebase-firestore.js";
import { getStorage, ref, uploadBytes, getDownloadURL } from "https://www.gstatic.com/firebasejs/10.8.0/firebase-storage.js";

// TODO: Remplacez par vos propres configurations Firebase issues de la console Firebase
const firebaseConfig = {
    apiKey: "VOTRE_API_KEY",
    authDomain: "VOTRE_AUTH_DOMAIN",
    projectId: "VOTRE_PROJECT_ID",
    storageBucket: "VOTRE_STORAGE_BUCKET",
    messagingSenderId: "VOTRE_MESSAGING_SENDER_ID",
    appId: "VOTRE_APP_ID"
};

const app = initializeApp(firebaseConfig);
const auth = getAuth(app);
const db = getFirestore(app);
const storage = getStorage(app);

let expensesCache = [];

// Gestion de la connexion
window.login = function() {
    const email = document.getElementById("email").value;
    const password = document.getElementById("password").value;
    signInWithEmailAndPassword(auth, email, password)
        .catch((error) => {
            document.getElementById("auth-error").innerText = "Erreur: " + error.message;
        });
}

window.logout = function() {
    signOut(auth);
}

onAuthStateChanged(auth, (user) => {
    if (user) {
        document.getElementById("auth-section").classList.add("hidden");
        document.getElementById("app-section").classList.remove("hidden");
        loadExpenses();
        document.getElementById("exp-date").valueAsDate = new Date();
    } else {
        document.getElementById("auth-section").classList.remove("hidden");
        document.getElementById("app-section").classList.add("hidden");
    }
});

// Calcul automatique TPS / TVQ
window.calculateTaxes = function() {
    const amount = parseFloat(document.getElementById("exp-amount").value) || 0;
    const isOtherProvince = document.getElementById("exp-other-province").checked;
    
    if (isOtherProvince) {
        document.getElementById("exp-tps").value = "0.00";
        document.getElementById("exp-tvq").value = "0.00";
    } else {
        // Au Québec, la TPS est 5% et la TVQ est 9.975% sur le montant avant taxes ou calcul inverse.
        // Formule standard : Total = HT + TPS + TVQ. Si on entre le montant TTC :
        // TPS = Total / 1.14975 * 0.05 ; TVQ = Total / 1.14975 * 0.09975
        const ht = amount / 1.14975;
        const tps = ht * 0.05;
        const tvq = ht * 0.09975;
        document.getElementById("exp-tps").value = tps.toFixed(2);
        document.getElementById("exp-tvq").value = tvq.toFixed(2);
    }
}

// Ajouter une dépense
window.addExpense = async function(event) {
    event.preventDefault();
    const date = document.getElementById("exp-date").value;
    const motif = document.getElementById("exp-motif").value;
    const total = parseFloat(document.getElementById("exp-amount").value);
    const tps = parseFloat(document.getElementById("exp-tps").value);
    const tvq = parseFloat(document.getElementById("exp-tvq").value);
    const photoFile = document.getElementById("exp-photo").files[0];

    try {
        // 1. Upload de la photo sur Firebase Storage
        const storageRef = ref(storage, `receipts/${Date.now()}_${photoFile.name}`);
        const snapshot = await uploadBytes(storageRef, photoFile);
        const photoUrl = await getDownloadURL(snapshot.ref);

        // 2. Enregistrement dans Firestore
        await addDoc(collection(db, "expenses"), {
            date,
            motif,
            total,
            tps,
            tvq,
            photoUrl,
            createdAt: new Date()
        });

        document.getElementById("expense-form").reset();
        document.getElementById("exp-date").valueAsDate = new Date();
        loadExpenses();
    } catch (error) {
        alert("Erreur lors de l'enregistrement : " + error.message);
    }
}

// Charger les dépenses
async function loadExpenses() {
    const tbody = document.getElementById("expenses-table-body");
    tbody.innerHTML = "<tr><td colspan='7' class='p-4 text-center text-gray-500'>Chargement...</td></tr>";
    
    const q = query(collection(db, "expenses"), orderBy("date", "desc"));
    const querySnapshot = await getDocs(q);
    
    expensesCache = [];
    tbody.innerHTML = "";

    querySnapshot.forEach((docSnap) => {
        const item = { id: docSnap.id, ...docSnap.data() };
        expensesCache.push(item);

        tbody.innerHTML += `
            <tr class="hover:bg-gray-50">
                <td class="p-3">${item.date}</td>
                <td class="p-3 font-medium text-gray-900">${item.motif}</td>
                <td class="p-3">${item.total.toFixed(2)} $</td>
                <td class="p-3">${item.tps.toFixed(2)} $</td>
                <td class="p-3">${item.tvq.toFixed(2)} $</td>
                <td class="p-3"><a href="${item.photoUrl}" target="_blank" class="text-blue-600 underline text-xs">Voir le reçu</a></td>
                <td class="p-3"><button onclick="deleteExpense('${item.id}')" class="text-red-500 hover:text-red-700 text-xs">Supprimer</button></td>
            </tr>
        `;
    });
}

// Supprimer
window.deleteExpense = async function(id) {
    if (confirm("Voulez-vous vraiment supprimer cette dépense ?")) {
        await deleteDoc(doc(db, "expenses", id));
        loadExpenses();
    }
}

// Export Excel organisé par événement avec liens vers les reçus
window.exportToExcel = function() {
    if (expensesCache.length === 0) {
        alert("Aucune donnée à exporter.");
        return;
    }

    // Transformation des données pour le fichier Excel
    const dataToExport = expensesCache.map(exp => ({
        "Date": exp.date,
        "Événement / Motif": exp.motif,
        "Montant Total ($)": exp.total,
        "TPS ($)": exp.tps,
        "TVQ ($)": exp.tvq,
        "Lien Reçu (Photo)": exp.photoUrl
    }));

    const worksheet = XLSX.utils.json_to_sheet(dataToExport);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Dépenses Travail");

    // Téléchargement du fichier .xlsx
    XLSX.writeFile(workbook, "Rapport_Depenses_Comptable.xlsx");
}
