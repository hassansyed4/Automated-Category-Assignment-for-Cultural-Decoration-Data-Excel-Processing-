import pandas as pd
import re

# File path to the Excel file
file_path = 'Categories v2 - Edited.xlsx'

# Load the Excel file
try:
    df = pd.read_excel(file_path, sheet_name='Sheet1')
except FileNotFoundError:
    print("Error: File not found.")
    exit()

# Utility to clean text
def clean_text(text):
    return text.lower() if isinstance(text, str) else ""

# Define your categories and rules 21 123 178 208 226 261 262 269 286 398
categories_keywords = {
    'Athletic': {
        'include': ['athlete', 'athletes', 'trainer', 'athletic', 'race', 'racing', 'horserace', 'chariot race'],
        'include_groups': [
                ['nude man', 'javelin'],
                ['man', 'diskos'],
                ['nude man', 'strigil'],
                ['youth', 'diskos'],  # Representing (Youth + Diskos)
        ],
        'exclude': ['Men', 'Youth', 'Group','Figure', 'Decorative', 'Animals', 'Animals-Mythical'],
    },
    'Warrior': {
        'include': ['pyrrhic','horsemen', 'arming', 'warrior', 'warriors', 'horseman', 'horsemen', 'warship', 'warships', 'chariot', 'chariots','archer'],
        'include_groups': [
                ['man', 'helmet'],
                ['man', 'spear'],
                ['man', 'sword'],
                ['man', 'shield'],
                ['youth', 'helmet'],
                ['youth', 'spear'],
                ['youths', 'spears'],
                ['youth', 'shield'],
                ['youth', 'sword'],
                ['rider', 'shield'],
                ['rider', 'sword'],
                ['rider', 'spear'],
        ],
        'exclude': ['Men', 'Youth','Figure','Group', 'Decorative', 'Animals', 'Animals-Mythical'],
    },
    'Warrior-Departing': {
        'include': ['warrior departing', 'warrior leaving', 'warrior exiting', 'warriors departing', 'warriors exiting', 'warriors leaving', 'chariot departing', 'chariot leaving', 'chariot exiting', 'chariots departing', 'chariots leaving', 'horseman departing', 'horseman leaving', 'horseman exiting', 'horsemen departing', 'horsemen leaving', 'horsemen exiting'],
        'include_groups': [
                ['warrior', 'departing'],
                ['warriors', 'departing'],
            ],
        'exclude': ['Gods','Ritual','Men', 'Youth', 'Warrior','Group','Women', 'Decorative', 'Animals', 'Animals-Mythical'],
    },
    'Symposion': {
        'include': ['symposion','symposium'],
        'exclude': ['Group', 'Music', 'Women', 'Men', 'Youth', 'Warrior', 'Figure', 'Decorative', 'Animals', 'Animals-Mythical'],
    },
    'Komos': {
        'include': ['komos'],
        'exclude': ['Group', 'Music', 'Women', 'Men', 'Figure','Youth', 'Decorative', 'Animals', 'Animals-Mythical'],
    },
    'Youth': {
        'include': ['boy','girl','youth', 'youths','child'],
        'exclude': ['Figure', 'Decorative', 'Animals', 'Animals-Mythical'],
    },
    'Men': {
        'include': ['huntsman','huntsmen','hunters','huntsmen', 'man', 'men', 'hunt', 'rider'],
        'include_groups': [
                ['man', 'youth'],
                ['man', 'youths'],
                ['men', 'youth'],
                ['men', 'youths'],
                ],
        'exclude': ['Youth','Women','Figure', 'Decorative', 'Animals', 'Animals-Mythical'],
    },
    'Women-Domestic': {
        'include': ['domestic'],
        'include_groups': [
            ['woman', 'chair', 'column'],
            ['women', 'chair', 'column'],
            ],
        'exclude': ['Women', 'Men', 'Youth', 'Group', 'Music','Gods', 'Decorative', 'Animals', 'Animals-Mythical'],
    },
    'Women': {
        'include': ['woman', 'women'],
        'exclude': ['Figure','Youth','Youths', 'Decorative', 'Animals', 'Animals-Mythical'],
    },
    'Group': {
        'include': ['group'],
        'include_groups': [
            ['figure', 'woman'],
            ['man', 'woman'],
            ['man', 'women'],
            ['men', 'women'],
            ['men', 'woman'],
            ['youth', 'woman'],
            ['youth', 'women'],
            ['youths', 'women'],
            ['youths', 'woman']
            # "Man + Woman" should all generate [Group]
                        ],
        'exclude': ['Women', 'Men', 'Youth', 'Warrior','Figure', 'Decorative', 'Animals', 'Animals-Mythical'], #'Group-Processional'
    },
    'Group-Wedding': {
        'include': ['wedding','wedded','bride','groom'], #added - 'bride','groom'
        'include_groups': [
            #['bride', 'groom'],
            ['bride', 'procession'],
            ['groom', 'procession']
            ],
        'exclude': ['Group', 'Man', 'Woman', 'Youth', 'Warrior', 'Music', 'Ritual', 'Gods', 'Group-Erotic','Dionysiac','Women', 'Decorative', 'Animals', 'Animals-Mythical'],
    },
    'Group-Erotic': {
        'include': ['courting', 'erotic', 'courtship'],
        'exclude': ['Women', 'Group','Men','Man', 'Woman', 'Youth', 'Warrior', 'Gods', 'Decorative', 'Animals', 'Animals-Mythical'],
    },
    'Figure': {
        'include': ['figure', 'figures', 'arm', 'leg', 'torso', 'head'], #381- Figure is coming bcz of HEAD, Gods for Athena.
        'exclude': ['Decorative'],
    },
    'Music': {
        'include': ['music contest','performance', 'instrument', 'instruments'],
        'include_groups': [
                ['kithara', 'judges'],
                ['kithara', 'wreath'],
                ['lyre', 'judges'],
                ['lyre', 'wreath'],
                ['pipes', 'judges'],
                ['pipes', 'wreath']
            ],
        'exclude': ['Decorative', 'Animals', 'Animals-Mythical'], #Group-Crowning
    },
    'Ritual': {
        'include': ['extispicy','procession','altar', 'shrine', 'nasikos','sacrifice'],
        'include_groups': [
                ['offering', 'stele'],
                ['phiale', 'oinochoe', 'woman'],
                ['cult', 'procession'],
                ['stele', 'offerings']
            ],
        'exclude': ['Figure', 'Group', 'Music', 'Group-Processional', 'Group-Offering', 'Youth', 'Men', 'Women', 'Decorative', 'Animals', 'Animals-Mythical'],
    },
    'Dionysiac': {
        'include': ['dionysos', 'satyr', 'satyrs', 'maenad', 'maenads', 'silen', 'thrysos', 'tymbales'],
        'exclude': ['Ritual', 'Symposion', 'Men', 'Youth', 'Women', 'Group', 'Music', 'Warrior', 'Figure', 'Decorative', 'Animals', 'Animals-Mythical'], # Remove - Gods
    },
    'Pursuit': {
        'include': ['chasing','pursuit', 'chase','pursuing'],
        'exclude': ['Group', 'Men', 'Youth', 'Women', 'Athletic', 'Figure', 'Decorative', 'Animals', 'Animals-Mythical'],
    },
    'Heroic': {
        'include': ['dioskouroi','perseus', 'chiron'],
        'exclude': ['Warrior', 'Warrior-Arming', 'Warrior-Departing', 'Men', 'Women', 'Youth', 'Group', 'Decorative', 'Animals', 'Animals-Mythical'],
    },
    'Heroic-Fight': {
        'include': ['fight', 'battle', 'combat', 'duel', 'warriors fighting', 'warrior fighting','grypomachy'],
        'exclude': ['Dionysiac', 'Figure', 'Warrior', 'Men', 'Youth', 'Women', 'Group', 'Pursuit', 'Decorative', 'Animals', 'Animals-Mythical'],
    },
    'Heroic-Amazonomachy': {
        'include': ['amazons fighting', 'amazonomachy','Amazon','Amazons'],
        'include_groups': [
                ['amazon', 'fight'],
                ['amazon', 'battle'],
                ['amazon', 'combat'],
                ['amazon', 'duel'],
                ['amazons', 'fight'],
                ['amazons', 'combat'],
                ['amazons', 'battle'],
            ],
        'exclude': ['Figure', 'Warrior', 'Heroic', 'Heroic-General Combat', 'Men', 'Youth', 'Women', 'Group', 'Pursuit', 'Heroic-Fight', 'Decorative', 'Animals', 'Animals-Mythical'],
    },
    'Heroic-Centauromachy': {
        'include': ['centaurs fighting', 'centauromachy'],
        'include_groups': [
                ['centaur', 'fight'],
                ['centaur', 'battle'],
                ['centaur', 'combat'],
                ['centaur', 'duel'],
                ['centaurs', 'battle'],
                ['centaurs', 'combat'],
                ['centaurs', 'fight'],
                ],
        'exclude': ['Warrior', 'Heroic', 'Heroic-General Combat', 'Men', 'Youth', 'Women', 'Group', 'Pursuit', 'Decorative', 'Animals', 'Animals-Mythical'],
    },
    'Heroic-Herakles': {
        'include': ['herakles', 'hercules', 'alcmene', 'iphicles', 'omphale', 'deianira', 'iolaus'],
        'exclude': ['Heroic-Centauromachy','Group', 'Men','Youth','Figure', 'Heroic-Fight', 'Figure', 'Heroic-Amazonomachy', 'Warrior', 'Warrior-Arming', 'Warrior-Departing', 'Heroic', 'Heroic-General Combat','Women', 'Decorative', 'Animals', 'Animals-Mythical'],
    },
    'Heroic-Theseus': {
        'include': ['theseus', 'aegeus', 'aethra', 'phaedra'],
        'exclude': ['Pursuit','Women','Figure', 'Group', 'Men', 'Youth', 'Warrior', 'Warrior-Arming', 'Warrior-Departing', 'Heroic', 'Heroic-General Combat', 'Decorative', 'Animals', 'Animals-Mythical'],
    },
    'Heroic-Troy': {
        'include': ['memnon','ilioupersis','aineias','troy', 'achilles', 'odysseus', 'agamemnon', 'ajax', 'hector', 'paris', 'helen', 'menelaus', 'priam', 'hecuba', 'cassandra', 'neoptolemus', 'laocoon'],
        'exclude': ['Youth','Funerary','Pursuit','Ritual', 'Men', 'Heroic-Fight', 'Group','Heroic-Amazonomachy','Women','Gods','Warrior', 'Warrior-Arming', 'Warrior-Departing', 'Heroic', 'Heroic-General Combat', 'Decorative', 'Animals', 'Animals-Mythical'],
    },
    'Myth': {
        'include': ['peleus','myth', 'return of hephaistos'],
        'include_groups': [
                ['anodos', 'figure'],
                ['anodos', 'man'],
                ['anodos', 'woman'],
                ],
        'exclude': ['Men','Women','Pursuit','Dionysiac', 'Gods', 'Decorative', 'Animals', 'Animals-Mythical'],
    },
    'Myth-Gigantomachy': {
        'include': ['gigantomachy', 'giants fighting'],
        'include_groups': [
                ['giant', 'combat'],
                ['giant', 'battle'],
                ['giant', 'fight'],
                ['giants', 'battle'],
                ['giants', 'combat'],
                ['giants', 'fight'],
                ],
        'exclude': ['Heroic-Fight','Men','Youth','Figure','Warrior','Dionysiac','Gods','Heroic-Herakles', 'Decorative', 'Animals', 'Animals-Mythical'],
    },
    'Myth-Triptoleomos': {
        'include': ['triptoleomos', 'triptolemos', 'triptolemus'],
        'include_groups': [
                ['triptoleomos', 'demeter'],
                ['triptoleomos', 'fire'],
                ['child', 'hearth']
        ],
        'exclude': ['Gods','Warrior','Ritual','Women','Decorative', 'Animals', 'Animals-Mythical'],
    },
    'Myth-Athena': {
        'include': [],
        'include_groups': [
                ['athena', 'olive tree'],
                ['athena', 'anodos'],
                ['athena', 'birth']
        ],
        'exclude': ['Men' ,'Figure','Gods','Warrior','Gods', 'Decorative', 'Animals', 'Animals-Mythical'],
    },
    'Gods': {
        'include': ['boreads','boread','winged youth','selene','erotes','eos','iris','hera','nereids','triton', 'tritons', 'assembly of gods', 'muse', 'muses', 'winged woman', 'winged man', 'winged figure','apollo', 'artemis', 'aphrodite', 'ares', 'zeus', 'poseidon', 'hades', 'demeter','gods','goddesses','god','goddess','Nike','hermes','athena','nymph','eros','nikai'],
        'exclude': ['Decorative', 'Animals', 'Animals-Mythical'],
    },
    'Funerary': {
        'include': ['funerary', 'funeral', 'mourning', 'grave', 'graves', 'tomb', 'naiskos'],
        'exclude': ['Youth', 'Group-Procession', 'Group', 'Men', 'Women', 'Youths', 'Figure', 'Gods', 'Ritual', 'Music', 'Warrior', 'Decorative', 'Animals', 'Animals-Mythical'],
    },
    'Theatrical': {
        'include': ['Theatrical','Actor','Actors'],
        'exclude': ['Youth','Men','Group','Decorative', 'Animals', 'Animals-Mythical'],
    },
    'Decorative': {
        'include': ['patterns','impressesd','diskos','ship','wreath','statue','grapevine','signature', 'basket', 'wineskin', 'post', 'pipecase', 'rosette','rosettes','eye', 'herm', 'column','star','signature','graffito','decorative', 'design', 'pattern', 'vine', 'vines','ivy', 'tendril', 'tendrils', 'floral', 'lotus', 'palmette', 'palmettes', 'palm', 'palms', 'sprig', 'gorgoneion', 'eyes', 'ships', 'plant','inscription','signature','whirligig'],
        'exclude': [],
    },
    'Animals' : {
        'include': ['dove','lioness','duck','ducks','rams','sheep','panthers', 'panther', 'owls' ,'hen', 'cocks', 'goats', 'goose', 'geese', 'animals', 'bulls','snake','bird', 'birds', 'owl', 'dog', 'dogs', 'horse', 'lion', 'ram', 'cat', 'animal','feline','dolphins','lions','horses','bull','swan', 'swans', 'panther', 'boar', 'goat','deer', 'eagle', 'eagles', 'dolphin','leopard','fish','cock','hare','mule'],
        'exclude': ['Decorative'],
    },
    'Animals-Mythical' : {
        'include': ['arimasp','hippalektryon','medusa','gorgon','gorgons','pygmies','minotaur','sirens', 'dwarf','chimaera','griffins','sphinx', 'sphinxes','griffin', 'gryphon','siren', 'sirens','centaur','centaurs','pegasos','pegasus','gorgon'],
        'exclude': ['Animals', 'Decorative'],
    }
}

# Classify plain text
def classify_text(text):
    text_lower = clean_text(text)
    matched_categories = []

    # Step 1: Loop through all categories to find possible matches
    for category, rules in categories_keywords.items():
        include_group_match = False
        include_match = False

        #Matching include_groups and keywords:

        # Check include_groups if defined
        include_groups = rules.get('include_groups', [])
        if include_groups:
            for group in include_groups:
                lowered_group = [term.lower() for term in group]  # <-- Create a new lowercase group | \b ensures whole-word matches
                if all(re.search(r'\b' + re.escape(term) + r'\b', text_lower) for term in lowered_group):
                    include_group_match = True
                    break

        # If no group matched, check simple includes
        if not include_group_match:
            include_match = any(
                re.search(r'\b' + re.escape(kw.lower()) + r'\b', text_lower)
                for kw in rules['include']
            )

        # If any match found, add category
        if include_group_match or include_match:
            matched_categories.append(category)

    # Step 2: Process exclusions
    final_categories = set(matched_categories)

    # For each category, remove its exclusions from the final_categories
    for cat in matched_categories:
        excludes = categories_keywords.get(cat, {}).get('exclude', [])
        for ex in excludes:
            final_categories.discard(ex)

    # If no categories remain after exclusions, categorize as "Uncategorized"
    return ", ".join(final_categories) if final_categories else "Uncategorized"

# Handle object areas like "Body:", "On rim:", etc.
def process_object_areas(description):
    output = []
    sections = re.split(r'\s*\|\s*', description)

    for section in sections:
        section = section.strip()
        match = re.match(r'^([\w\s,]+:)\s*(.*)', section)  # e.g., "Body:" or "On rim:"

        if match:
            object_area = match.group(1).strip()
            section_text = match.group(2).strip()
            categories = classify_text(section_text) if section_text else 'Uncategorized'
            output.append(f"{object_area} {categories}")
        else:
            categories = classify_text(section) if section else 'Uncategorized'
            output.append(categories)

    return ' | '.join(output)

# Apply the categorization to 'Decoration'
if 'Decoration' in df.columns:
    df['Categories'] = df['Decoration'].astype(str).apply(process_object_areas)
else:
    print("'Decoration' column not found!")

# Save to the same file (or you can rename the output file)
df.to_excel(file_path, index=False)

print("Classification completed. Results saved in 'Categories v2 - Edited - Output.xlsx'")
