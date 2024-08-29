import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd

# Definizione dei settori e delle relative descrizioni
settori = [
    {"name": "Healthcare", "description": "Oral and Maxillofacial Surgery, Medical, Otolaryngology clinic, Mental  clinic, Home  care service, medical, Pharmaceutical, Social and s, Social and Healthcares, Home Healthcare care service, not-for-profit Healthcare system, Mental Healthcare clinic, Nutrition, Clinic, Pharmaceuticals, Rehabilitation, Dental Services, Public Health, Personal care, Health, Healthcare services, Health Care Provider, Helth Care, Medical Equipments, Medical Devices, Heath Care, Biotechnology, Medical Equipment, Health Care, Health Care Providers, Hospitals, Clinics, Oral and maxillofacial surgery,Oral and Maxillofacial Surgery, Medical Supplies, Mental Healthcare, Social and Healthcare, biomedical life sciences, Pharmaceutics"},
    {"name": "Financial", "description": "BANK, Title company, Holding Company, Credit union, Financial agency, financial advisor, Market infrastructures, IT/Financial, Tax, Chartered Accountants, Accounting, Financial advisor, Fintech, Financial services, Banking institutions, Finance, Financial organizations, Specialty REITs, Insurance Brokers, Private Equity, Equity Investment Instruments, Full Line Insurance, Banking, Mortgage Finance, Residential REITs, Specialty Finance, Investment Services, Financial Services, Banking, Insurance, Investment Companies, Banks, Financial Administration"},
    {"name": "Energy", "description": "Oil & natural gas, Energia, hydrocarbon exploration, Electrical, Hydrocarbon exploration, Oil and Gas Services, Solar Energy, Oil, Utilities, Renewable Energy, Oil and Gas, Water distribution and supply, Gas, Electricity, Renewable energies, Utility, Plumbing & Electrical, Gas Distribution, Alternative Fuels, Renewable Energ, Integrated Oil & Gas, Multi-utilities, Conventional Electricity, Oil & Gas, Energy, Renewable Energy Equipment, Oil Equipment & Services, Oil and gas companies, alternative energy producers and suppliers"},
    {"name": "Industrial", "description": "Manufacturing, manufacturer, Industrial Manufacturing, Pharmacy and drugs Manufacturing, Welding supply, Lighting manufacturer, chemical company, Sheet metal contractor, Manufacturing    , Boat builders, Manufacturing, HVAC, manufacture, Indaustrial Manufacturing, Manufacturer, Industrial Supplies, Lumber, Plastics, Mechanical Services, Metals, Aerospace, Machinery, Architecture, Chemicals, Lighting, Auto Repair, HVAC Manufacturing, Engineering, Automotive Services, Steel Manufacturing, Automotive, Manufacturing, Foundry, Mining, Engineering and automation, Heavy industries, Robotics, Farming, Infrastructures, Gold Mining, Automobiles & Parts, Chemical, Electrical Equipment, Industrial Products, Water Treatment, Industrial Suppliers, Equipment, Building Materials, Waste Disposal, Forestry, Packaging, Materials, Precision Engineering, Aerospace & Defense, Iron & Steel, Aluminum, Auto Parts, Steel, Waste Treatment, Tobacco, Paper, Waste & Disposal Services, Exploration & Production, Paper & Wood Products, Industrial Services, Basic Resources, Commodity Chemicals, Diversified Industrials, Containers & Packaging, Automobiles, Chemicals, Fishing & Plantations, Specialty Chemicals, Chemical Processing, Engineering and Manufacturing Companies, Waste Treatment, Industrial Machinery, Industrial Goods & Services"},
    {"name": "Construction", "description": "Contractor, Mechanical contractor, builders, Real estate consultant, Real estate agency, Heating contractor, Construccion, Interior Design, Road construction company, Construction company, General contractor, Plumbing, Home Improvement, Architecture, Climate, Engineering & Constructions, Heavy Constructions, Construction & Engineering, Construction & Materials, Enginering & Construction, Constructions, Building Materials & Fixtures, Diversified Industrials, Heavy Construction, Real Estate, Real estate, Home Construction, real estate, Real Estate Holding & Development, Real Estate Services, Building Materials & Fixtures"},
    {"name": "Technology", "description": "business process automation, Software company, Computer consultant, Metrology, Information Technology, Internet Service Provider, IT Consulting, Audio Equipment, IT Solutions, Cybersecurity, Software, Electronics, IT Services, Development, Hosting service providers, Information Technologies Consulting, Cloud service providers, Digital infrastructures, Internet Service providers, High-tech, Technologies, Technology Services, Hosting Provider, Software & Computer Services, Electrical Components, Cloud Provider, Internet Services, Telecommunication, Electrical Components & Equipment, Telecommunications Equipment, Technology, Internet, Telecommunicaitons, Electronic Equipment, Computer Hardware, Telecommunications, Semiconductors, Electronic Office Equipment, hardware companies, Computer Services"},
    {"name": "Education", "description": "university, school, School, Training, Academic Institutions, Universities, Schools, Education Services, Public and private universities and colleges, training and development companies"},
    {"name": "Services", "description": "Business Services, Digital service provider, service engineering solutions provider, Digital printing service, security agency, Employment agency, Consultant, consulting, Fire protection service, Environmental, Administration, International Trade, Security, Inc., business, Professional services, Advertising services, Information services, Business services, HR Services, Email Services, HVAC Services, Cleaning Services, Business Service, Corporate office, Business, Human Resources, International Organization and Consulting, Dry Cleaning, Linen Services, Recruitment, Project Management, Postal Services, Consultancy, Waste Management, Conglomerate, Certification, Quality Assurance, Auction, Safety, Staffing, Consulting, Security Services, Self Storage, Accounts Management, Wealth Management, Global services, Human Services, Engineering consulting, Insurance services, Information Technologies Consulting, Legal consulting, Attorney, Labor Organization, Process Outsourcing, Asset Management, Consumer Support Services, Legal services, IT Services, Accountanting, Asset Managers, Business Training & Employment Agencies, Business Support Services, Professional Services, Legal Services, Accounting and Consulting Firms, Recreational Services, Recreational Products"},
    {"name": "Entertainment", "description": "Opera, sport, Art, Event Management, Sports, Animation, Culture, Culture and entertainment, Entertainment industry, Movie production, Sports, Gaming, Casinos, Broadcasting, Gambling, Think tanks"},
    {"name": "Transportation", "description": "Warehouse, Shipping service, Bus charter, Freight forwarding service, automotive, transport, Freight, Traffic Management, Mobility, Transport, Marine Equipment, Shipping, Marine, Import/Export, Aviation, Air Mobility, Helicopter Maintenance, Air Transportation, Logistics, Flight, Avionics, Rail transport, Air transport, Maritime transport, Road transport, Logistics, Airlines, Transportation Services, Automotive, Railroads, Commercial Vehicles & Trucks, Delivery Services, Railroad, Trucking, Transporter, Travel"},
    {"name": "Communication", "description": "service telefonica, Comunication, Telecommunications, Communications, Marketing, Marketing & Advertising, Advertising, Publishing, Newspapers, book publishers, public relations and advertising agencies"},
    {"name": "Consumer", "description": "food, Foods & Pharmacy, Custom label printer, Wholesale bakery, Brewing company, Food products supplier, agency specialising in lighting, Winery, Agriculture , Agriculture, Aquaculture, Textile, Food, Restaurant supply, Bakery, Beauty, Gardening, Furniture, Beverage, Trade, Packaging, Trading, Agriculture, Paper Products, Fragrances, Textiles, Fashion, Embroidery, Culture and Apparel, Luxury, Sports, Apparel, Shops, Pharmacy and drugs manufacturing, House & Office Products, Nondurable Household Products, Consumer Goods, Household Goods, Consumer Staple Products, Apparel & Textile, Furniship, Textile & Apparel, Home & Office Products, Furniture, Household Products, Cosmetics, Footwear, Furnishing, Consumer Electronics, Consumer Services, Toys, Clothing & Accessories, ersonal Products, Personal & Household Goods, Durable Household Products, Jewelry, Wholesale, Physical Security, Furnishings, Specialized Consumer Services, Manufacturers and distributors of consumer products"},
    {"name": "Media", "description": "Blogging, News, Music Distribution, Broadcasting, Printing, Publishing and Media, Religious Broadcasting, Advertising, Medias and audiovisual, Media Agencies, Broadcasting & Entertainment, Television, Satellite, Social Media, Internet"},
    {"name": "Hospitality", "description": "travel, hotel, Tourism, Food Service, Food Production, Food Industry, Travel and tourism industry, Lodging industry, Hotel, Food and drinks businesses, Beverage, Distillers & Vintners, Beverages, Food Products, s Pizza, Restaurants & Bars, Hotels, Restaurants, Cruise Lines, Travel & Leisure, Leisure Services, Leisure Facilities, Food & Beverage, Agriculture and agribusiness"},
    {"name": "Retail", "description": "Retail , fashion, Retail , Retail (distribution), Retail , Medical supply store, Sign shop, Vitamin & supplements store, Furniture store, Building materials store, store online, store, Sales, Jewelry, Sporting Goods, E-commerce, Distribution, Drug Retailers, Broadline Retailers, Food Retailers & Wholesalers, Retail Apparel, Specialty Retailers, Home Improvement Retailers, Apparel Retailers, Brick and mortar and e-commerce"},
    {"name": "Research", "description": "Research, Healthcare research, Research & Development, Market Research, Think Tanks, Research and Development"},
    {"name": "Public", "description": "law , law, government , law firm, Non-Governmental Organizations, Non-Governmental Organizations , Fighting poverty, not-for-profit  system, Religious organization, force of order, City government office, Law firm, Lawyer, Non-Governmental Organizations, Religion, Defence, Government, Diplomacy, Sheriff's department, Defense ministries, Legislative branch, Non-Governmental Organizations, Central administration and government, Charity, Humanitarian, SLU, Nonprofit, LLP, Legal, PLLC, Law Enforcement, Non-profit, Local Government, Union, Non-profit, Law, Lawyers, Dissidents, Critical Sectors, Defense research and development, Ministries of foreign affairs, Political parties, Public sector, Judicial power (justice), Critical Industries, Defense industry, International organizations, Charity associations, Defense ministries (including the military), Defense, Church, Non-Governmental Organizations (NGOs), Legislative branch (parliamentary chambers), Civil society, entral administration and government, Citizens, Government and administrations, Local administrations, Defense, Governmenta Agencies, Governent Agencies, Military, Government Agencies, Federal"},
    {"name": "n.d.", "description": "Unknown, unknown"}
]


def carica_e_modifica_excel():
    root = tk.Tk()
    root.withdraw()  # Nasconde la finestra principale di Tkinter

    # Chiede all'utente di selezionare un file Excel
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if not file_path:  # Se l'utente non seleziona un file, esce dalla funzione
        messagebox.showinfo("Info", "Caricamento annullato.")
        return root.destroy()

    try:
        df = pd.read_excel(file_path)

        # Verifica che la colonna 'Victim sectors' esista
        if 'Victim sectors' not in df.columns:
            messagebox.showerror("Errore", "La colonna 'Victim sectors' non è presente nel file Excel.")
            return root.destroy()

        # Modifica il campo 'Victim sectors' in base ai settori definiti
        for settore in settori:
            for descrizione in settore["description"].split(", "):
                df['Victim sectors'] = df['Victim sectors'].replace(descrizione, settore["name"])

        # Chiede all'utente dove salvare il file Excel modificato
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if save_path:
            df.to_excel(save_path, index=False)
            messagebox.showinfo("Successo", "Il file Excel è stato modificato e salvato con successo.")
        else:
            messagebox.showinfo("Info", "Salvataggio annullato.")

    except Exception as e:
        messagebox.showerror("Errore", f"Si è verificato un errore: {e}")

    finally:
        root.destroy()

if __name__ == "__main__":
    carica_e_modifica_excel()