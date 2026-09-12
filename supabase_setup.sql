-- Table des utilisateurs (app_users pour éviter les conflits avec auth.users)
CREATE TABLE public.app_users (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    nom TEXT UNIQUE NOT NULL,
    pin TEXT NOT NULL,
    role TEXT NOT NULL DEFAULT 'user',
    chantiers TEXT NOT NULL DEFAULT '' -- '*' pour admin, ou liste 'Chantier 1, Chantier 2'
);

-- Table des chantiers
CREATE TABLE public.chantiers (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    nom TEXT UNIQUE NOT NULL
);

-- Table des saisies (charges, produits, reglements)
CREATE TABLE public.saisies (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    chantier TEXT NOT NULL,
    date DATE NOT NULL DEFAULT CURRENT_DATE,
    libelle TEXT NOT NULL,
    charges NUMERIC DEFAULT 0,
    produits NUMERIC DEFAULT 0,
    reglement NUMERIC DEFAULT 0,
    obs TEXT DEFAULT '',
    ref_vers TEXT DEFAULT '',
    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

-- Table des versements
CREATE TABLE public.versements (
    id TEXT PRIMARY KEY, -- ex: 'V-001'
    date DATE NOT NULL DEFAULT CURRENT_DATE,
    chantier TEXT NOT NULL,
    montant NUMERIC NOT NULL DEFAULT 0,
    obs TEXT DEFAULT '',
    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

-- Table des libellés (Listes)
CREATE TABLE public.listes (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    libelle TEXT NOT NULL,
    chantier_id TEXT DEFAULT '',
    cp TEXT DEFAULT ''
);

-- Désactivation RLS (Row Level Security) pour un démarrage simple, 
-- puisque l'app gère elle-même l'authentification avec nom/PIN.
ALTER TABLE public.app_users DISABLE ROW LEVEL SECURITY;
ALTER TABLE public.chantiers DISABLE ROW LEVEL SECURITY;
ALTER TABLE public.saisies DISABLE ROW LEVEL SECURITY;
ALTER TABLE public.versements DISABLE ROW LEVEL SECURITY;
ALTER TABLE public.listes DISABLE ROW LEVEL SECURITY;

-- Insertion de l'utilisateur Admin par défaut
INSERT INTO public.app_users (nom, pin, role, chantiers) 
VALUES ('Admin Principal', '0000', 'admin', '*')
ON CONFLICT (nom) DO NOTHING;

-- Insertion d'un chantier de test
INSERT INTO public.chantiers (nom) 
VALUES ('Chantier Test')
ON CONFLICT (nom) DO NOTHING;

-- Insertion des libellés par défaut
INSERT INTO public.listes (libelle) VALUES 
('Achat matériaux'),
('Location engin'),
('Carburant'),
('Main d''œuvre'),
('Transport'),
('Paiement client'),
('Avance client'),
('Facture fournisseur'),
('Frais divers');
