import { useEffect, useMemo, useRef, useState } from "react";
import productsData from "./data/products.json";
import productImages from "./data/product-images.json";
import {
  downloadCartExcel,
  printCartAsPdf,
  type CartExportRow,
} from "./cartExports";

type Product = {
  sku: string;
  description: string;
  capacity: string;
  price: string;
  barcode: string;
};

type CartItem = {
  product: Product;
  quantity: string;
};

type SortOrder = "catalog" | "name" | "price-asc" | "price-desc" | "sku";
type PriceFilter = "all" | "priced" | "unpriced";

type CategoryGroupId =
  | "cloths"
  | "scrubbing"
  | "kitchen"
  | "bags-laundry"
  | "cleaning"
  | "home"
  | "health";

type ProductCategoryId =
  | "wet-cloths"
  | "microfiber"
  | "floor-cloths"
  | "multi-cloths"
  | "scrub-pads"
  | "scourer-rolls"
  | "steel-wool"
  | "sponges"
  | "tablecloths"
  | "gloves"
  | "trash-bags"
  | "cotton-buds"
  | "clothespins"
  | "toothpicks"
  | "skewers"
  | "candles"
  | "foil"
  | "kitchen-products"
  | "storage"
  | "laundry"
  | "air-fresheners"
  | "brushes-brooms"
  | "mops"
  | "handles"
  | "dustpans"
  | "silicone"
  | "toilet-tools"
  | "buckets-bins"
  | "cleaning-accessories"
  | "aprons-hats"
  | "stirrers"
  | "dispensers"
  | "pharmacy"
  | "other";

type CategoryId = "all" | CategoryGroupId | ProductCategoryId;

type CategoryGroup = {
  id: CategoryGroupId;
  label: string;
  categories: ProductCategoryId[];
};

type FeaturedBrand = {
  id: string;
  name: string;
  category: ProductCategoryId;
  sku: string;
  tagline: string;
  title: string;
  description: string;
  action: string;
  theme: "gorilla" | "clip-n-grip";
};

const products: Product[] = productsData;
const productImageBySku: Record<string, string> = productImages;
const PAGE_SIZE = 24;
const VAT_RATE = 0.18;
const featuredBrands: FeaturedBrand[] = [
  {
    id: "gorilla",
    name: "GORILLA",
    category: "trash-bags",
    sku: "1425",
    tagline: "מותג שקיות האשפה",
    title: "חוזק שאפשר לסמוך עליו",
    description: "שקיות Gorilla למגוון צרכים בבית ובעסק.",
    action: "לצפייה בשקיות אשפה",
    theme: "gorilla",
  },
  {
    id: "clip-n-grip",
    name: "CLIP ’N GRIP",
    category: "clothespins",
    sku: "0978",
    tagline: "מותג אטבי הכביסה",
    title: "אחיזה חזקה, יום אחרי יום",
    description: "הכירו את אטבי הכביסה של Clip ’N Grip.",
    action: "לצפייה באטבי כביסה",
    theme: "clip-n-grip",
  },
];
const sortOptions: { value: SortOrder; label: string }[] = [
  { value: "catalog", label: "סדר הקטלוג" },
  { value: "name", label: "שם מוצר (א–ת)" },
  { value: "price-asc", label: "מחיר: מהנמוך לגבוה" },
  { value: "price-desc", label: "מחיר: מהגבוה לנמוך" },
  { value: "sku", label: "מק״ט" },
];
const priceFilterOptions: { value: PriceFilter; label: string }[] = [
  { value: "all", label: "כל המחירים" },
  { value: "priced", label: "מחיר זמין" },
  { value: "unpriced", label: "מחיר חסר" },
];

const categoryLabels: Record<ProductCategoryId, string> = {
  "wet-cloths": "מטליות לחות",
  microfiber: "מטליות מיקרופייבר",
  "floor-cloths": "סחבות לרצפה",
  "multi-cloths": "מטליות רב שימושיות",
  "scrub-pads": "כריות קרצוף כלים",
  "scourer-rolls": "סרטי קרצוף (סקוטש)",
  "steel-wool": "קרצוף מקצועי",
  sponges: "ספוגי ניקוי ורחצה",
  tablecloths: "מפות שולחן",
  gloves: "כפפות",
  "trash-bags": "שקיות אשפה",
  "cotton-buds": "צמרוני כותנה",
  clothespins: "אטבי כביסה",
  toothpicks: "קיסמי שיניים",
  skewers: "שיפודים ומדליקי גחלים",
  candles: "נרות",
  foil: "רדידי אלומיניום",
  "kitchen-products": "מוצרים למטבח",
  storage: "שקיות ואחסון",
  laundry: "אביזרי כביסה",
  "air-fresheners": "מבשמי אוויר",
  "brushes-brooms": "מברשות ומטאטאים",
  mops: "מגבים",
  handles: "מקלות מטאטא ומגב",
  dustpans: "יעים",
  silicone: "מוצרי סיליקון",
  "toilet-tools": "מברשות אסלה ופומפות",
  "buckets-bins": "דליים ופחים",
  "cleaning-accessories": "אביזרי ניקיון",
  "aprons-hats": "סינרים וכובעים",
  stirrers: "בוחשנים לשתייה",
  dispensers: "מתקני סבון ונייר",
  pharmacy: "מוצרי פארם",
  other: "מוצרים נוספים",
};

const categoryGroups: CategoryGroup[] = [
  {
    id: "cloths",
    label: "מטליות וניגוב",
    categories: ["multi-cloths", "floor-cloths", "microfiber", "wet-cloths"],
  },
  {
    id: "scrubbing",
    label: "ספוגים וקרצוף",
    categories: ["scrub-pads", "scourer-rolls", "steel-wool", "sponges"],
  },
  {
    id: "kitchen",
    label: "מטבח ואירוח",
    categories: [
      "tablecloths",
      "gloves",
      "cotton-buds",
      "toothpicks",
      "skewers",
      "candles",
      "foil",
      "kitchen-products",
    ],
  },
  {
    id: "bags-laundry",
    label: "שקיות וכביסה",
    categories: ["trash-bags", "storage", "clothespins", "laundry"],
  },
  {
    id: "cleaning",
    label: "כלי ניקיון",
    categories: [
      "brushes-brooms",
      "mops",
      "handles",
      "dustpans",
      "silicone",
      "toilet-tools",
      "buckets-bins",
      "cleaning-accessories",
    ],
  },
  {
    id: "home",
    label: "מוצרים לבית",
    categories: ["air-fresheners", "aprons-hats", "stirrers", "dispensers", "other"],
  },
  {
    id: "health",
    label: "פארם",
    categories: ["pharmacy"],
  },
];

const categoryGroupDescriptions: Record<CategoryGroupId, string> = {
  cloths: "מטליות, סחבות ומוצרי ניגוב לבית ולעסק",
  scrubbing: "ספוגים, כריות קרצוף ומוצרי ניקוי",
  kitchen: "מוצרים שימושיים למטבח ולאירוח",
  "bags-laundry": "שקיות, אטבים ואביזרים לכביסה",
  cleaning: "כלי ניקיון, מגבים ומברשות",
  home: "מוצרים משלימים לבית וליום־יום",
  health: "מוצרי פארם ועזרה ראשונה",
};

const categoryRules: { id: ProductCategoryId; pattern: RegExp }[] = [
  { id: "wet-cloths", pattern: /מגבונים|מגבון|לחות לרצפה|לח לרצפה|נייר טואלט לח/ },
  { id: "microfiber", pattern: /מיקרופייבר/ },
  { id: "floor-cloths", pattern: /סחב|סחבות/ },
  { id: "multi-cloths", pattern: /מטלית|מטליות|מגבת|מגבות|גליל מטליות/ },
  { id: "scrub-pads", pattern: /כרית|כריות/ },
  { id: "scourer-rolls", pattern: /סקוטש|סרטי? קרצוף/ },
  { id: "steel-wool", pattern: /צמר פלדה|ברזליות|ממרקות/ },
  { id: "sponges", pattern: /ספוג|ספונג/ },
  { id: "tablecloths", pattern: /מפה|מפת|מפות/ },
  { id: "gloves", pattern: /כפפה|כפפות/ },
  { id: "trash-bags", pattern: /אשפה|שקיות זבל|שקית זבל/ },
  { id: "cotton-buds", pattern: /צמרוני|מקלוני אוזניים/ },
  { id: "clothespins", pattern: /אטב|אטבים/ },
  { id: "toothpicks", pattern: /קיסמ|קיסם/ },
  { id: "skewers", pattern: /שיפוד|מדליק גחלים/ },
  { id: "candles", pattern: /נרות|נר / },
  { id: "foil", pattern: /רדיד|אלומיניום|נייר כסף/ },
  {
    id: "kitchen-products",
    pattern: /מטבח|קופס|ואקום|וואקום|ניילון נצמד|נייר אפייה|אפייה|מפיות|כפיות|מצע לייבוש|PVC|EVA|סנדוויץ|הקפאה/,
  },
  {
    id: "storage",
    pattern: /שקית|שקיות|כיסוי|קליפסים|אחסון/,
  },
  { id: "laundry", pattern: /חבל כביסה|קולב|כביסה/ },
  {
    id: "toilet-tools",
    pattern: /אסלה|פומפה|שירותים|שרותים|משתנות/,
  },
  {
    id: "air-fresheners",
    pattern: /מבשם|מבשמי|מפיץ ריח|ריחני|מטהר אויר|מטהר אוויר/,
  },
  { id: "brushes-brooms", pattern: /מברשת|מברשות|מטאטא/ },
  { id: "mops", pattern: /מגב/ },
  { id: "handles", pattern: /מקל/ },
  { id: "dustpans", pattern: /יעה|כף אשפה/ },
  { id: "silicone", pattern: /סיליקון|VIVIAN|ויויאן/i },
  { id: "buckets-bins", pattern: /דלי|דליים|פח|פחים|סל נייר|סל כביסה|סל / },
  { id: "cleaning-accessories", pattern: /מרסס|תרסיס|בקבוק מרסס/ },
  { id: "aprons-hats", pattern: /סינר|כובע/ },
  { id: "stirrers", pattern: /בוחש|מערבל/ },
  { id: "dispensers", pattern: /מתקן|דיספנסר|נייר טואלט|נייר מגבת|ממחטות|טישיו/ },
  {
    id: "pharmacy",
    pattern: /Farma|פלסטר|אלכוהול|חיטוי|מסכה|תחבושת|פד |אגד |מד חום|עזרה ראשונה|צמר גפן/i,
  },
];

function getCategory(description: string): ProductCategoryId {
  const category = categoryRules.find(({ pattern }) => pattern.test(description));
  return category?.id ?? "other";
}

function getCategoryGroup(category: ProductCategoryId): CategoryGroupId {
  const group = categoryGroups.find(({ categories }) =>
    categories.includes(category),
  );
  if (!group) {
    throw new Error(`Missing category group for ${category}`);
  }
  return group.id;
}

function isCategoryGroup(category: CategoryId): category is CategoryGroupId {
  return categoryGroups.some((group) => group.id === category);
}

const routableCategoryIds = new Set<string>([
  ...categoryGroups.map(({ id }) => id),
  ...categoryRules.map(({ id }) => id),
  "other",
]);

function getCategoryFromUrl(): CategoryId | null {
  const category = new URLSearchParams(window.location.search).get("category");
  return category && routableCategoryIds.has(category)
    ? (category as CategoryId)
    : null;
}

function getOrderMultiple(capacity: string) {
  const amounts = capacity.match(/\d+/g)?.map(Number);
  return amounts?.length
    ? amounts.reduce((total, amount) => total * amount, 1)
    : 1;
}

function getNumericPrice(price: string) {
  if (!price) {
    return null;
  }

  const amount = Number(price.replace(",", "."));
  return Number.isFinite(amount) ? amount : null;
}

function formatCurrency(amount: number) {
  return new Intl.NumberFormat("he-IL", {
    style: "currency",
    currency: "ILS",
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  }).format(amount);
}

const variantAttributeTerms = [
  "כחול שמיים",
  "כחול כהה",
  "כהה",
  "בהיר",
  "ורוד בייבי",
  "מגוון צבעים",
  "פריחת הדובדבן",
  "בית מלון יוקרתי",
  "מרכך כביסה",
  "מלון יוקרה",
  "פצ'ולי",
  "פצולי",
  "לבנדר",
  "וניל",
  "לימון",
  "פרחים",
  "מרענן",
  "שחורות",
  "שחורים",
  "שחורה",
  "שחור",
  "לבנות",
  "לבנים",
  "לבנה",
  "לבן",
  "כחולות",
  "כחולים",
  "כחולה",
  "כחול",
  "צהובות",
  "צהובים",
  "צהובה",
  "צהוב",
  "אדומות",
  "אדומים",
  "אדומה",
  "אדום",
  "ירוקות",
  "ירוקים",
  "ירוקה",
  "ירוק",
  "ורודות",
  "ורודים",
  "ורודה",
  "ורוד",
  "סגולות",
  "סגולים",
  "סגולה",
  "סגול",
  "אפורות",
  "אפורים",
  "אפורה",
  "אפור",
  "כתומות",
  "כתומים",
  "כתומה",
  "כתום",
  "חומות",
  "חומים",
  "חומה",
  "חום",
  "זהובה",
  "זהב",
  "כסופה",
  "כסוף",
  "כסף",
  "נחושת",
  "צבעוניות",
  "צבעוניים",
  "צבעונית",
  "צבעוני",
];

function getVariantGroupKey(description: string) {
  let key = description
    .toLocaleLowerCase("he")
    .replace(/\d+\s*(?:[x×]\s*\d+)+/giu, " ")
    .replace(
      /(?:כ\s*-\s*)?\d+(?:[.,]\d+)?\s*(?:מטר|מ['״]|ס['״]?מ|מ['״]?ל|מל|ליטר|גרם|קילו|ק['״]?ג|יחידות|יח['״]?)/giu,
      " ",
    )
    .replace(/(?:מידה\s*)?\b(?:XXL|XL|XS|S|M|L)\b/giu, " ")
    .replace(/(?:בניחוח|ניחוח)\s*/giu, " ")
    .replace(/בלבד/giu, " ");

  for (const term of variantAttributeTerms) {
    key = key.replace(new RegExp(term, "giu"), " ");
  }

  key = key.replace(/[^A-Za-z0-9\u0590-\u05FF]+/g, " ").replace(/\s+/g, " ").trim();
  return key.split(" ").filter(Boolean).length >= 2 ? key : null;
}

const productsByVariantKey = new Map<string, Product[]>();
for (const product of products) {
  const key = getVariantGroupKey(product.description);
  if (!key) {
    continue;
  }

  const variants = productsByVariantKey.get(key) ?? [];
  variants.push(product);
  productsByVariantKey.set(key, variants);
}

const variantsBySku = new Map<string, Product[]>();
for (const variants of productsByVariantKey.values()) {
  if (variants.length < 2) {
    continue;
  }

  for (const product of variants) {
    variantsBySku.set(product.sku, variants);
  }
}

function SearchIcon() {
  return (
    <svg
      aria-hidden="true"
      className="search-icon"
      viewBox="0 0 24 24"
      fill="none"
    >
      <circle cx="10.8" cy="10.8" r="6.8" />
      <path d="m16 16 4.2 4.2" />
    </svg>
  );
}

function ProductCard({
  product,
  index,
  quantityInCart,
  onAdd,
  onAdjust,
  onDetailsClick,
}: {
  product: Product;
  index: number;
  quantityInCart: number;
  onAdd: (product: Product) => void;
  onAdjust: (sku: string, direction: -1 | 1) => void;
  onDetailsClick: (product: Product) => void;
}) {
  const category = getCategory(product.description);
  const number = String(index + 1).padStart(2, "0");
  const orderMultiple = getOrderMultiple(product.capacity);
  const image = productImageBySku[product.sku];
  const unitPrice = getNumericPrice(product.price);

  return (
    <article className="product-card">
      <div className={`product-visual visual-${category}${image ? " has-photo" : ""}`}>
        {image ? (
          <button
            className="product-photo-button"
            type="button"
            aria-label={`פתיחת פרטי מוצר ${product.description}`}
            onClick={() => onDetailsClick(product)}
          >
            <img
              className="product-photo"
              src={image}
              alt=""
              loading="lazy"
            />
            <span className="photo-zoom-hint" aria-hidden="true">
              <svg viewBox="0 0 24 24" fill="none">
                <circle cx="10.8" cy="10.8" r="6.8" />
                <path d="m16 16 4.2 4.2M10.8 7.5v6.6M7.5 10.8h6.6" />
              </svg>
            </span>
          </button>
        ) : (
          <>
            <span className="visual-orbit orbit-one" />
            <span className="visual-orbit orbit-two" />
            <span className="visual-center" aria-hidden="true">
              <svg viewBox="0 0 48 48" fill="none">
                <path d="M24 5 28.7 18.3 42 23l-13.3 4.7L24 41l-4.7-13.3L6 23l13.3-4.7L24 5Z" />
              </svg>
            </span>
            <span className="visual-index">{number}</span>
            <span className="visual-category">{categoryLabels[category]}</span>
          </>
        )}
      </div>
      <div className="product-copy">
        <div className="product-topline">
          <span className="product-sku">מק״ט {product.sku || "—"}</span>
          {unitPrice !== null ? (
            <span className="product-price">
              <span className="meta-label">מחיר ליחידה:</span>{" "}
              {formatCurrency(unitPrice)}
            </span>
          ) : (
            <span className="product-price product-price-unavailable">
              מחיר לא זמין
            </span>
          )}
        </div>
        <h2>
          <button
            className="product-title-button"
            type="button"
            onClick={() => onDetailsClick(product)}
          >
            {product.description || "מוצר שייני"}
          </button>
        </h2>
        <div className="product-meta">
          <span>
            <span className="meta-label">אריזה:</span>{" "}
            {product.capacity || "לא צוינה"}
          </span>
          {product.barcode && (
            <span className="barcode">
              <span className="meta-label">ברקוד:</span> {product.barcode}
            </span>
          )}
        </div>
        {quantityInCart > 0 ? (
          <div className="product-cart-controls">
            <div className="product-quantity-control">
              <button
                type="button"
                aria-label={`הפחתת כמות ${product.description}${quantityInCart === orderMultiple ? ", הסרת הפריט מהעגלה" : ""}`}
                title={
                  quantityInCart === orderMultiple
                    ? "הסרת הפריט מהעגלה"
                    : "הפחתת אריזה אחת"
                }
                onClick={() => onAdjust(product.sku, -1)}
              >
                −
              </button>
              <span className="product-quantity" aria-live="polite">
                {quantityInCart}
              </span>
              <button
                type="button"
                aria-label={`הוספת כמות ${product.description}`}
                title="הוספת אריזה אחת"
                onClick={() => onAdjust(product.sku, 1)}
              >
                +
              </button>
            </div>
            <span className="add-multiple">בעגלה · כפולות של {orderMultiple}</span>
          </div>
        ) : (
          <button
            className="add-to-cart"
            type="button"
            onClick={() => onAdd(product)}
          >
            <span>+ הוספה לעגלה</span>
            <span className="add-multiple">כפולות של {orderMultiple}</span>
          </button>
        )}
      </div>
    </article>
  );
}

function App() {
  const [query, setQuery] = useState("");
  const [searchSuggestionsOpen, setSearchSuggestionsOpen] = useState(false);
  const [activeSearchSuggestion, setActiveSearchSuggestion] = useState(-1);
  const [pageCategory, setPageCategory] = useState<CategoryId | null>(
    getCategoryFromUrl,
  );
  const [priceFilter, setPriceFilter] = useState<PriceFilter>("all");
  const [mobileFiltersOpen, setMobileFiltersOpen] = useState(false);
  const [sortOrder, setSortOrder] = useState<SortOrder>("catalog");
  const [visibleCount, setVisibleCount] = useState(PAGE_SIZE);
  const [cart, setCart] = useState<CartItem[]>([]);
  const [cartOpen, setCartOpen] = useState(false);
  const [selectedProduct, setSelectedProduct] = useState<Product | null>(null);
  const [accessibilityOpen, setAccessibilityOpen] = useState(false);
  const [highContrastEnabled, setHighContrastEnabled] = useState(false);
  const [underlinedLinksEnabled, setUnderlinedLinksEnabled] = useState(false);
  const [accessibilityTextScale, setAccessibilityTextScale] = useState(1);
  const [strongFocusEnabled, setStrongFocusEnabled] = useState(false);
  const [reducedMotionEnabled, setReducedMotionEnabled] = useState(false);
  const [expandedSpacingEnabled, setExpandedSpacingEnabled] = useState(false);
  const [readableFontEnabled, setReadableFontEnabled] = useState(false);
  const [activeBrandIndex, setActiveBrandIndex] = useState(0);
  const [brandCarouselHovered, setBrandCarouselHovered] = useState(false);
  const [brandCarouselFocused, setBrandCarouselFocused] = useState(false);
  const siteHeaderRef = useRef<HTMLElement | null>(null);

  useEffect(() => {
    if (!cartOpen && !selectedProduct && !accessibilityOpen) {
      return;
    }

    function closeOnEscape(event: KeyboardEvent) {
      if (event.key === "Escape") {
        setCartOpen(false);
        setSelectedProduct(null);
        setAccessibilityOpen(false);
      }
    }

    document.addEventListener("keydown", closeOnEscape);
    return () => document.removeEventListener("keydown", closeOnEscape);
  }, [accessibilityOpen, cartOpen, selectedProduct]);

  useEffect(() => {
    document.documentElement.classList.toggle(
      "accessibility-reduced-motion",
      reducedMotionEnabled,
    );
    return () =>
      document.documentElement.classList.remove(
        "accessibility-reduced-motion",
      );
  }, [reducedMotionEnabled]);

  useEffect(() => {
    const header = siteHeaderRef.current;
    if (!header) {
      return;
    }

    const updateHeaderHeight = () => {
      document.documentElement.style.setProperty(
        "--site-header-height",
        `${header.offsetHeight}px`,
      );
    };
    const observer = new ResizeObserver(updateHeaderHeight);
    observer.observe(header);
    updateHeaderHeight();

    return () => {
      observer.disconnect();
      document.documentElement.style.removeProperty("--site-header-height");
    };
  }, []);

  useEffect(() => {
    function syncCategoryFromUrl() {
      setPageCategory(getCategoryFromUrl());
      setQuery("");
      setSearchSuggestionsOpen(false);
      setActiveSearchSuggestion(-1);
      setPriceFilter("all");
      setSortOrder("catalog");
      setVisibleCount(PAGE_SIZE);
      setMobileFiltersOpen(false);
      window.scrollTo({ top: 0, behavior: "auto" });
    }

    window.addEventListener("popstate", syncCategoryFromUrl);
    return () => window.removeEventListener("popstate", syncCategoryFromUrl);
  }, []);

  useEffect(() => {
    if (
      reducedMotionEnabled ||
      brandCarouselHovered ||
      brandCarouselFocused ||
      pageCategory !== null ||
      Boolean(query.trim()) ||
      window.matchMedia("(prefers-reduced-motion: reduce)").matches
    ) {
      return;
    }

    const intervalId = window.setInterval(() => {
      setActiveBrandIndex((index) => (index + 1) % featuredBrands.length);
    }, 6000);

    return () => window.clearInterval(intervalId);
  }, [
    brandCarouselFocused,
    brandCarouselHovered,
    pageCategory,
    query,
    reducedMotionEnabled,
  ]);

  const categoryCounts = useMemo(() => {
    const categoryIds: CategoryId[] = [
      "all",
      ...categoryGroups.map(({ id }) => id),
      ...categoryRules.map(({ id }) => id),
      "other",
    ];
    const counts = Object.fromEntries(
      categoryIds.map((category) => [category, 0]),
    ) as Record<CategoryId, number>;

    products.forEach((product) => {
      const category = getCategory(product.description);
      counts[category] += 1;
      counts[getCategoryGroup(category)] += 1;
    });
    counts.all = products.length;
    return counts;
  }, []);

  const selectedCategory = pageCategory ?? "all";
  const showProductPage = pageCategory !== null || Boolean(query.trim());
  const categoryLandingCards = useMemo(
    () =>
      categoryGroups
        .filter((group) => categoryCounts[group.id] > 0)
        .map((group) => {
          const representativeProduct = products.find(
            (product) =>
              getCategoryGroup(getCategory(product.description)) === group.id &&
              Boolean(productImageBySku[product.sku]),
          );
          return {
            group,
            count: categoryCounts[group.id],
            description: categoryGroupDescriptions[group.id],
            image: representativeProduct
              ? productImageBySku[representativeProduct.sku]
              : null,
          };
        }),
    [categoryCounts],
  );

  const categoryPageTitle =
    selectedCategory === "all"
      ? "תוצאות חיפוש"
      : isCategoryGroup(selectedCategory)
        ? categoryGroups.find(({ id }) => id === selectedCategory)?.label ?? ""
        : categoryLabels[selectedCategory];

  const selectedGroupId =
    selectedCategory === "all"
      ? null
      : isCategoryGroup(selectedCategory)
        ? selectedCategory
        : getCategoryGroup(selectedCategory);
  const selectedGroup = categoryGroups.find(
    ({ id }) => id === selectedGroupId,
  );

  const filteredProducts = useMemo(() => {
    const normalizedQuery = query.trim().toLocaleLowerCase("he");

    return products.filter((product) => {
      const productCategory = getCategory(product.description);
      const matchesCategory =
        selectedCategory === "all" ||
        (isCategoryGroup(selectedCategory)
          ? getCategoryGroup(productCategory) === selectedCategory
          : productCategory === selectedCategory);
      const searchableText = [
        product.description,
        product.sku,
        product.capacity,
        product.barcode,
      ]
        .join(" ")
        .toLocaleLowerCase("he");
      const productPrice = getNumericPrice(product.price);
      const matchesPrice =
        priceFilter === "all" ||
        (priceFilter === "priced"
          ? productPrice !== null
          : productPrice === null);

      return (
        matchesCategory &&
        matchesPrice &&
        searchableText.includes(normalizedQuery)
      );
    });
  }, [priceFilter, query, selectedCategory]);

  const sortedProducts = useMemo(() => {
    if (sortOrder === "catalog") {
      return filteredProducts;
    }

    return [...filteredProducts].sort((first, second) => {
      if (sortOrder === "name") {
        return first.description.localeCompare(second.description, "he");
      }
      if (sortOrder === "sku") {
        return first.sku.localeCompare(second.sku, "he", { numeric: true });
      }

      const firstPrice = getNumericPrice(first.price);
      const secondPrice = getNumericPrice(second.price);
      if (firstPrice === null || secondPrice === null) {
        if (firstPrice === secondPrice) {
          return 0;
        }
        return firstPrice === null ? 1 : -1;
      }

      return sortOrder === "price-asc"
        ? firstPrice - secondPrice
        : secondPrice - firstPrice;
    });
  }, [filteredProducts, sortOrder]);

  const searchSuggestions = useMemo(() => {
    const normalizedQuery = query.trim().toLocaleLowerCase("he");
    if (!normalizedQuery) {
      return [];
    }

    return sortedProducts
      .map((product, index) => {
        const sku = product.sku.toLocaleLowerCase("he");
        const description = product.description.toLocaleLowerCase("he");
        const barcode = product.barcode.toLocaleLowerCase("he");
        const capacity = product.capacity.toLocaleLowerCase("he");
        const score =
          sku === normalizedQuery
            ? 0
            : sku.startsWith(normalizedQuery)
              ? 1
              : description.startsWith(normalizedQuery)
                ? 2
                : description.includes(normalizedQuery)
                  ? 3
                  : barcode.includes(normalizedQuery) ||
                      capacity.includes(normalizedQuery)
                    ? 4
                    : 5;

        return { product, index, score };
      })
      .sort((first, second) => first.score - second.score || first.index - second.index)
      .slice(0, 6)
      .map(({ product }) => product);
  }, [query, sortedProducts]);

  const visibleProducts = sortedProducts.slice(0, visibleCount);
  const selectedProductVariants = selectedProduct
    ? variantsBySku.get(selectedProduct.sku) ?? []
    : [];
  const selectedProductQuantity = selectedProduct
    ? Number(
        cart.find((item) => item.product.sku === selectedProduct.sku)?.quantity,
      ) || 0
    : 0;
  const selectedProductPrice = selectedProduct
    ? getNumericPrice(selectedProduct.price)
    : null;
  const showSearchSuggestions =
    searchSuggestionsOpen && Boolean(query.trim());
  const activeFilterCount =
    Number(selectedCategory !== "all") + Number(priceFilter !== "all");
  const cartCount = cart.reduce(
    (total, item) => total + (Number(item.quantity) || 0),
    0,
  );
  const cartRows = useMemo<CartExportRow[]>(
    () =>
      cart.map(({ product, quantity }) => {
        const unitPrice = getNumericPrice(product.price);
        const count = Number(quantity);
        const multiple = getOrderMultiple(product.capacity);
        const validQuantity =
          Number.isInteger(count) &&
          count >= multiple &&
          count % multiple === 0;

        return {
          sku: product.sku,
          description: product.description,
          barcode: product.barcode,
          capacity: product.capacity,
          quantity: count,
          unitPrice,
          lineTotal:
            unitPrice !== null && validQuantity
              ? Math.round(unitPrice * count * 100) / 100
              : null,
        };
      }),
    [cart],
  );
  const cartSubtotal =
    Math.round(
      cartRows.reduce((total, row) => total + (row.lineTotal ?? 0), 0) * 100,
    ) / 100;
  const cartVat = Math.round(cartSubtotal * VAT_RATE * 100) / 100;
  const cartTotalWithVat =
    Math.round((cartSubtotal + cartVat) * 100) / 100;
  const hasInvalidQuantity = cart.some(({ product, quantity }) => {
    const amount = Number(quantity);
    const multiple = getOrderMultiple(product.capacity);
    return (
      !Number.isInteger(amount) ||
      amount < multiple ||
      amount % multiple !== 0
    );
  });

  function updateQuery(value: string) {
    setQuery(value);
    setVisibleCount(PAGE_SIZE);
    setSearchSuggestionsOpen(Boolean(value.trim()));
    setActiveSearchSuggestion(-1);
  }

  function openProductDetails(product: Product) {
    setSelectedProduct(product);
    setSearchSuggestionsOpen(false);
    setActiveSearchSuggestion(-1);
  }

  function selectSearchSuggestion(product: Product) {
    updateQuery(product.sku || product.description);
    openProductDetails(product);
  }

  function navigateToCategory(category: CategoryId | null) {
    const url = new URL(window.location.href);
    url.hash = "";
    if (category) {
      url.searchParams.set("category", category);
    } else {
      url.searchParams.delete("category");
    }

    if (url.href !== window.location.href) {
      window.history.pushState(null, "", url);
    }
    setPageCategory(category);
    updateQuery("");
    setPriceFilter("all");
    setSortOrder("catalog");
    setMobileFiltersOpen(false);
    window.scrollTo({
      top: 0,
      behavior: reducedMotionEnabled ? "auto" : "smooth",
    });
  }

  function updateCategory(category: CategoryId) {
    navigateToCategory(category === "all" ? null : category);
  }

  function updatePriceFilter(value: PriceFilter) {
    setPriceFilter(value);
    setVisibleCount(PAGE_SIZE);
  }

  function updateSortOrder(value: SortOrder) {
    setSortOrder(value);
    setVisibleCount(PAGE_SIZE);
  }

  function showBrandCategory(category: ProductCategoryId) {
    navigateToCategory(category);
  }

  function addToCart(product: Product) {
    const multiple = getOrderMultiple(product.capacity);
    setCart((items) => {
      const existing = items.find(
        (item) => item.product.sku === product.sku,
      );

      if (existing) {
        return items.map((item) =>
          item.product.sku === product.sku
            ? {
                ...item,
                quantity: String(
                  (Number(item.quantity) || 0) + multiple,
                ),
              }
            : item,
        );
      }

      return [...items, { product, quantity: String(multiple) }];
    });
  }

  function updateCartQuantity(sku: string, quantity: string) {
    setCart((items) =>
      items.map((item) =>
        item.product.sku === sku ? { ...item, quantity } : item,
      ),
    );
  }

  function adjustCartQuantity(sku: string, direction: -1 | 1) {
    setCart((items) =>
      items.flatMap((item) => {
        if (item.product.sku !== sku) {
          return item;
        }

        const multiple = getOrderMultiple(item.product.capacity);
        const rawQuantity = Number(item.quantity);
        const quantity = Number.isFinite(rawQuantity) ? rawQuantity : 0;
        const valid =
          Number.isInteger(quantity) &&
          quantity >= multiple &&
          quantity % multiple === 0;
        const nextQuantity = valid
          ? quantity + direction * multiple
          : direction === 1
            ? Math.ceil(quantity / multiple) * multiple || multiple
            : Math.floor(quantity / multiple) * multiple;

        if (nextQuantity <= 0) {
          return [];
        }

        return {
          ...item,
          quantity: String(Math.max(multiple, nextQuantity)),
        };
      }),
    );
  }

  function removeFromCart(sku: string) {
    setCart((items) => items.filter((item) => item.product.sku !== sku));
  }

  const featuredBrand = featuredBrands[activeBrandIndex];

  return (
    <main
      className={[
        highContrastEnabled ? "accessibility-high-contrast" : "",
        underlinedLinksEnabled ? "accessibility-underlined-links" : "",
        accessibilityTextScale > 1 ? "accessibility-large-text" : "",
        accessibilityTextScale > 1
          ? `accessibility-text-scale-${Math.round(accessibilityTextScale * 100)}`
          : "",
        strongFocusEnabled ? "accessibility-strong-focus" : "",
        expandedSpacingEnabled ? "accessibility-expanded-spacing" : "",
        readableFontEnabled ? "accessibility-readable-font" : "",
      ]
        .filter(Boolean)
        .join(" ")}
    >
      <header className="site-header" ref={siteHeaderRef}>
        <div className="site-header-top-row">
          <a
            className="brand"
            href={showProductPage ? "/" : "#top"}
            aria-label="שייני - חזרה לראש העמוד"
            id="top"
            onClick={(event) => {
              if (showProductPage) {
                event.preventDefault();
                navigateToCategory(null);
              }
            }}
          >
            <img className="brand-logo" src="/shiny-logo.png" alt="Shiny" />
            <span className="brand-divider" />
            <span className="brand-caption">מוצרי ניקיון</span>
          </a>
          <div className="header-actions">
            <button
              className="header-link login-placeholder"
              type="button"
              disabled
              title="ההתחברות תתווסף בהמשך"
              aria-label="התחברות לאתר, בקרוב"
            >
              <span>התחברות</span>
              <span className="login-coming-soon">בקרוב</span>
            </button>
            <button
              className="cart-toggle"
              type="button"
              onClick={() => setCartOpen(true)}
              aria-label={`פתיחת עגלת קניות, ${cartCount} פריטים`}
            >
              <svg viewBox="0 0 24 24" fill="none" aria-hidden="true">
                <path d="M3 4h2l2.2 11.1a2 2 0 0 0 2 1.6h8.5a2 2 0 0 0 1.9-1.4L21 9H6" />
                <circle cx="10" cy="20" r="1" />
                <circle cx="18" cy="20" r="1" />
              </svg>
              <span>עגלה</span>
              <span className="cart-badge">{cartCount}</span>
            </button>
          </div>
        </div>
        <div className="catalog-sticky-bar">
          <div className="catalog-controls">
            <div className="catalog-control-fields">
              <div className="search-wrap">
                <SearchIcon />
                <input
                  id="catalog-search"
                  aria-label="חיפוש מוצרים"
                  type="text"
                  role="combobox"
                  aria-autocomplete="list"
                  aria-expanded={showSearchSuggestions}
                  aria-controls={
                    showSearchSuggestions
                      ? "catalog-search-suggestions"
                      : undefined
                  }
                  aria-haspopup="listbox"
                  aria-activedescendant={
                    activeSearchSuggestion >= 0 &&
                    activeSearchSuggestion < searchSuggestions.length
                      ? `search-suggestion-${activeSearchSuggestion}`
                      : undefined
                  }
                  value={query}
                  onChange={(event) => updateQuery(event.target.value)}
                  onFocus={() => {
                    if (query.trim()) {
                      setSearchSuggestionsOpen(true);
                    }
                  }}
                  onBlur={(event) => {
                    const nextTarget = event.relatedTarget;
                    if (
                      !(nextTarget instanceof Node) ||
                      !event.currentTarget.contains(nextTarget)
                    ) {
                      setSearchSuggestionsOpen(false);
                      setActiveSearchSuggestion(-1);
                    }
                  }}
                  onKeyDown={(event) => {
                    if (event.key === "Escape" && showSearchSuggestions) {
                      event.preventDefault();
                      setSearchSuggestionsOpen(false);
                      setActiveSearchSuggestion(-1);
                    } else if (searchSuggestions.length > 0) {
                      if (event.key === "ArrowDown") {
                        event.preventDefault();
                        setSearchSuggestionsOpen(true);
                        setActiveSearchSuggestion((active) =>
                          active >= searchSuggestions.length - 1
                            ? 0
                            : active + 1,
                        );
                      } else if (event.key === "ArrowUp") {
                        event.preventDefault();
                        setSearchSuggestionsOpen(true);
                        setActiveSearchSuggestion((active) =>
                          active <= 0
                            ? searchSuggestions.length - 1
                            : active - 1,
                        );
                      } else if (
                        event.key === "Enter" &&
                        activeSearchSuggestion >= 0
                      ) {
                        event.preventDefault();
                        const suggestion =
                          searchSuggestions[activeSearchSuggestion];
                        if (suggestion) {
                          selectSearchSuggestion(suggestion);
                        }
                      }
                    }
                  }}
                  placeholder="שם מוצר, מותג, ניחוח, גודל, מק״ט או ברקוד"
                />
                {query && (
                  <button
                    className="clear-search"
                    type="button"
                    onClick={() => updateQuery("")}
                    aria-label="ניקוי החיפוש"
                  >
                    ×
                  </button>
                )}
                {showSearchSuggestions && (
                  <div
                    className="search-suggestions"
                    id="catalog-search-suggestions"
                    role="listbox"
                    aria-label="הצעות למוצרים"
                  >
                    {searchSuggestions.length > 0 ? (
                      searchSuggestions.map((product, index) => {
                        const price = getNumericPrice(product.price);
                        return (
                          <button
                            className={`search-suggestion${activeSearchSuggestion === index ? " active" : ""}`}
                            id={`search-suggestion-${index}`}
                            key={product.sku}
                            type="button"
                            role="option"
                            aria-selected={activeSearchSuggestion === index}
                            onMouseDown={(event) => event.preventDefault()}
                            onMouseEnter={() =>
                              setActiveSearchSuggestion(index)
                            }
                            onClick={() => selectSearchSuggestion(product)}
                          >
                            <span className="search-suggestion-copy">
                              <strong>{product.description}</strong>
                              <small>מק״ט {product.sku}</small>
                            </span>
                            <span className="search-suggestion-price">
                              {price === null
                                ? "מחיר לא זמין"
                                : formatCurrency(price)}
                            </span>
                          </button>
                        );
                      })
                    ) : (
                      <p className="search-suggestions-empty">
                        לא נמצאו הצעות מתאימות
                      </p>
                    )}
                  </div>
                )}
              </div>
              {showProductPage && (
                <label className="sort-control">
                  <span>מיון</span>
                  <select
                    aria-label="מיון מוצרים"
                    value={sortOrder}
                    onChange={(event) => {
                      const option = sortOptions.find(
                        ({ value }) => value === event.target.value,
                      );
                      if (option) {
                        updateSortOrder(option.value);
                      }
                    }}
                  >
                    {sortOptions.map((option) => (
                      <option key={option.value} value={option.value}>
                        {option.label}
                      </option>
                    ))}
                  </select>
                </label>
              )}
            </div>
            {showProductPage && (
              <button
                className="mobile-filter-toggle"
                type="button"
                aria-expanded={mobileFiltersOpen}
                aria-controls="catalog-extra-filters"
                onClick={() => setMobileFiltersOpen((open) => !open)}
              >
                <span>מסננים</span>
                {activeFilterCount > 0 && (
                  <span className="filter-toggle-count">{activeFilterCount}</span>
                )}
                <span aria-hidden="true">{mobileFiltersOpen ? "−" : "+"}</span>
              </button>
            )}
          </div>

          <nav
            className={`category-navigation${showProductPage ? "" : " categories-only"}`}
            id="catalog-category-filters"
            aria-label={
              showProductPage
                ? "סינון לפי קטגוריה ומחיר"
                : "ניווט לקטגוריה"
            }
          >
            <p className="category-navigation-label">קטגוריות</p>
            <div className="category-list category-groups">
              <button
                className={`category-chip${selectedCategory === "all" ? " active" : ""}`}
                type="button"
                aria-pressed={selectedCategory === "all"}
                onClick={() => updateCategory("all")}
              >
                כל הקטגוריות
              </button>
              {categoryGroups
                .filter((group) => categoryCounts[group.id] > 0)
                .map((group) => (
                  <button
                    className={`category-chip${selectedGroupId === group.id ? " active" : ""}`}
                    key={group.id}
                    type="button"
                    aria-pressed={selectedGroupId === group.id}
                    onClick={() => updateCategory(group.id)}
                  >
                    {group.label}
                    <span>{categoryCounts[group.id]}</span>
                  </button>
                ))}
            </div>

            <div
              className={`category-extra-filters${mobileFiltersOpen ? " mobile-open" : ""}`}
              id="catalog-extra-filters"
            >
              {selectedGroup && (
                <div className="subcategory-panel">
                  <p className="subcategory-heading">
                    בתוך {selectedGroup.label}
                  </p>
                  <div className="category-list category-subcategories">
                    <button
                      className={`category-chip${selectedCategory === selectedGroup.id ? " active" : ""}`}
                      type="button"
                      aria-pressed={selectedCategory === selectedGroup.id}
                      onClick={() => updateCategory(selectedGroup.id)}
                    >
                      כל הקטגוריה
                      <span>{categoryCounts[selectedGroup.id]}</span>
                    </button>
                    {selectedGroup.categories
                      .filter((category) => categoryCounts[category] > 0)
                      .map((category) => (
                        <button
                          className={`category-chip${selectedCategory === category ? " active" : ""}`}
                          key={category}
                          type="button"
                          aria-pressed={selectedCategory === category}
                          onClick={() => updateCategory(category)}
                        >
                          {categoryLabels[category]}
                          <span>{categoryCounts[category]}</span>
                        </button>
                      ))}
                  </div>
                </div>
              )}

              {showProductPage && (
                <div className="price-filter-group">
                  <p className="subcategory-heading">זמינות מחיר</p>
                  <div
                    className="category-list"
                    role="group"
                    aria-label="סינון לפי זמינות מחיר"
                  >
                    {priceFilterOptions.map((option) => (
                      <button
                        className={`category-chip${priceFilter === option.value ? " active" : ""}`}
                        key={option.value}
                        type="button"
                        aria-pressed={priceFilter === option.value}
                        onClick={() => updatePriceFilter(option.value)}
                      >
                        {option.label}
                      </button>
                    ))}
                  </div>
                </div>
              )}
            </div>
          </nav>
        </div>
      </header>

      {!showProductPage && featuredBrand && (
        <section className="featured-brands-section" aria-label="מותגים מובילים">
          <div
            className={`brand-promo-card brand-promo-card-${featuredBrand.theme}`}
            id="featured-brand-promo"
            role="group"
            aria-label={featuredBrand.name}
            aria-roledescription="קרוסלת מותגים"
            onMouseEnter={() => setBrandCarouselHovered(true)}
            onMouseLeave={() => setBrandCarouselHovered(false)}
            onFocus={() => setBrandCarouselFocused(true)}
            onBlur={(event) => {
              const nextTarget = event.relatedTarget;
              if (
                !(nextTarget instanceof Node) ||
                !event.currentTarget.contains(nextTarget)
              ) {
                setBrandCarouselFocused(false);
              }
            }}
          >
            <div className="brand-promo-copy" key={`${featuredBrand.id}-copy`}>
              <p className="brand-promo-tagline">{featuredBrand.tagline}</p>
              <p className="brand-promo-wordmark">{featuredBrand.name}</p>
              <h2>{featuredBrand.title}</h2>
              <p className="brand-promo-description">
                {featuredBrand.description}
              </p>
              <button
                className="brand-promo-action"
                type="button"
                onClick={() => showBrandCategory(featuredBrand.category)}
              >
                {featuredBrand.action}
                <span aria-hidden="true">←</span>
              </button>
            </div>
            <div
              className="brand-promo-art"
              key={`${featuredBrand.id}-art`}
              aria-hidden="true"
            >
              <span className="brand-promo-watermark">
                {featuredBrand.name}
              </span>
              <img
                src={productImageBySku[featuredBrand.sku]}
                alt=""
                loading="lazy"
              />
            </div>
            <div className="brand-promo-controls">
              <button
                className="brand-promo-arrow"
                type="button"
                aria-label="המותג הקודם"
                aria-controls="featured-brand-promo"
                onClick={() =>
                  setActiveBrandIndex(
                    (index) =>
                      (index - 1 + featuredBrands.length) %
                      featuredBrands.length,
                  )
                }
              >
                <span aria-hidden="true">→</span>
              </button>
              <div
                className="brand-promo-dots"
                role="group"
                aria-label="בחירת מותג"
              >
                {featuredBrands.map((brand, index) => (
                  <button
                    key={brand.id}
                    className={`brand-promo-dot${index === activeBrandIndex ? " active" : ""}`}
                    type="button"
                    aria-label={`הצגת ${brand.name}`}
                    aria-pressed={index === activeBrandIndex}
                    aria-controls="featured-brand-promo"
                    onClick={() => setActiveBrandIndex(index)}
                  />
                ))}
              </div>
              <button
                className="brand-promo-arrow"
                type="button"
                aria-label="המותג הבא"
                aria-controls="featured-brand-promo"
                onClick={() =>
                  setActiveBrandIndex(
                    (index) => (index + 1) % featuredBrands.length,
                  )
                }
              >
                <span aria-hidden="true">←</span>
              </button>
            </div>
          </div>
        </section>
      )}

      {!showProductPage && (
        <section className="category-landing" id="categories">
          <div className="category-landing-heading">
            <div>
              <p className="eyebrow section-eyebrow">הקטלוג שלנו</p>
              <h1>מוצאים את מה שצריך לפי קטגוריה</h1>
              <p>בחרו תחום כדי לראות את המוצרים הרלוונטיים.</p>
            </div>
          </div>
          <div className="category-tile-grid">
            {categoryLandingCards.map(
              ({ group, count, description, image }) => (
                <button
                  className="category-tile"
                  key={group.id}
                  type="button"
                  onClick={() => navigateToCategory(group.id)}
                  aria-label={`פתיחת קטגוריית ${group.label}, ${count} מוצרים`}
                >
                  <span className="category-tile-image" aria-hidden="true">
                    {image ? (
                      <img src={image} alt="" loading="lazy" />
                    ) : (
                      <span className="category-tile-placeholder">✳</span>
                    )}
                  </span>
                  <span className="category-tile-copy">
                    <span className="category-tile-title">{group.label}</span>
                    <span className="category-tile-description">
                      {description}
                    </span>
                    <span className="category-tile-footer">
                      <span>{count} מוצרים</span>
                      <span>
                        לצפייה בקטגוריה <span aria-hidden="true">←</span>
                      </span>
                    </span>
                  </span>
                </button>
              ),
            )}
          </div>
        </section>
      )}

      {showProductPage && (
        <section className="catalog-section" id="catalog">
          <div className="catalog-heading">
            <div>
              <p className="eyebrow section-eyebrow">
                {query.trim() ? "תוצאות חיפוש" : "הקטלוג שלנו"}
              </p>
              <h1>{categoryPageTitle}</h1>
            </div>
            <div className="catalog-total">
              <span className="total-number">{filteredProducts.length}</span>
              <span>מוצרים</span>
            </div>
          </div>

          {visibleProducts.length > 0 ? (
            <div className="product-grid">
              {visibleProducts.map((product, index) => (
                <ProductCard
                  key={`${product.sku}-${product.barcode}-${index}`}
                  product={product}
                  index={sortedProducts.indexOf(product)}
                  quantityInCart={
                    Number(
                      cart.find((item) => item.product.sku === product.sku)
                        ?.quantity,
                    ) || 0
                  }
                  onAdd={addToCart}
                  onAdjust={adjustCartQuantity}
                  onDetailsClick={openProductDetails}
                />
              ))}
            </div>
          ) : (
            <div className="empty-state">
              <span className="empty-icon" aria-hidden="true">
                ✳
              </span>
              <h3>לא מצאנו מוצרים מתאימים</h3>
              <p>נסו לחפש במילים אחרות או לבחור קטגוריה אחרת.</p>
              <button type="button" onClick={() => updateCategory("all")}>
                חזרה לקטגוריות
              </button>
            </div>
          )}

          {visibleCount < sortedProducts.length && (
            <div className="load-more-wrap">
              <button
                className="load-more"
                type="button"
                onClick={() => setVisibleCount((count) => count + PAGE_SIZE)}
              >
                עוד מוצרים
                <span aria-hidden="true">↓</span>
              </button>
            </div>
          )}
        </section>
      )}

      <footer className="site-footer">
        <a className="footer-brand" href="#top" aria-label="שייני - חזרה לראש העמוד">
          <img src="/shiny-logo.png" alt="Shiny" />
        </a>
        <span>עושים מקום לניקיון טוב יותר.</span>
        <a href="#top">בחזרה למעלה ↑</a>
      </footer>

      {!cartOpen && !selectedProduct && (
        <div className="accessibility-widget" dir="rtl">
          {accessibilityOpen && (
            <section
              className="accessibility-panel"
              id="accessibility-panel"
              aria-labelledby="accessibility-title"
            >
              <header className="accessibility-panel-header">
                <h2 id="accessibility-title">אפשרויות נגישות</h2>
                <button
                  className="accessibility-panel-close"
                  type="button"
                  aria-label="סגירת תפריט הנגישות"
                  onClick={() => setAccessibilityOpen(false)}
                >
                  ×
                </button>
              </header>
              <button
                className="accessibility-option"
                type="button"
                aria-pressed={highContrastEnabled}
                onClick={() => setHighContrastEnabled((enabled) => !enabled)}
              >
                <span>ניגודיות גבוהה</span>
                <span className="accessibility-option-state">
                  {highContrastEnabled ? "פעיל" : "כבוי"}
                </span>
              </button>
              <button
                className="accessibility-option"
                type="button"
                aria-pressed={underlinedLinksEnabled}
                onClick={() =>
                  setUnderlinedLinksEnabled((enabled) => !enabled)
                }
              >
                <span>הדגשת קישורים בקו תחתון</span>
                <span className="accessibility-option-state">
                  {underlinedLinksEnabled ? "פעיל" : "כבוי"}
                </span>
              </button>
              <div className="accessibility-text-controls">
                <span>גודל טקסט</span>
                <div>
                  <button
                    type="button"
                    aria-label="הקטנת טקסט"
                    title="הקטנת טקסט"
                    disabled={accessibilityTextScale <= 1}
                    onClick={() =>
                      setAccessibilityTextScale((scale) =>
                        Math.max(1, Math.round((scale - 0.1) * 10) / 10),
                      )
                    }
                  >
                    A−
                  </button>
                  <span aria-live="polite">
                    {Math.round(accessibilityTextScale * 100)}%
                  </span>
                  <button
                    type="button"
                    aria-label="הגדלת טקסט"
                    title="הגדלת טקסט"
                    disabled={accessibilityTextScale >= 1.2}
                    onClick={() =>
                      setAccessibilityTextScale((scale) =>
                        Math.min(1.2, Math.round((scale + 0.1) * 10) / 10),
                      )
                    }
                  >
                    A+
                  </button>
                </div>
              </div>
              <button
                className="accessibility-option"
                type="button"
                aria-pressed={strongFocusEnabled}
                onClick={() => setStrongFocusEnabled((enabled) => !enabled)}
              >
                <span>הדגשת מיקוד מקלדת</span>
                <span className="accessibility-option-state">
                  {strongFocusEnabled ? "פעיל" : "כבוי"}
                </span>
              </button>
              <button
                className="accessibility-option"
                type="button"
                aria-pressed={reducedMotionEnabled}
                onClick={() => setReducedMotionEnabled((enabled) => !enabled)}
              >
                <span>הפחתת אנימציות</span>
                <span className="accessibility-option-state">
                  {reducedMotionEnabled ? "פעיל" : "כבוי"}
                </span>
              </button>
              <button
                className="accessibility-option"
                type="button"
                aria-pressed={expandedSpacingEnabled}
                onClick={() => setExpandedSpacingEnabled((enabled) => !enabled)}
              >
                <span>ריווח מוגדל בטקסט</span>
                <span className="accessibility-option-state">
                  {expandedSpacingEnabled ? "פעיל" : "כבוי"}
                </span>
              </button>
              <button
                className="accessibility-option"
                type="button"
                aria-pressed={readableFontEnabled}
                onClick={() => setReadableFontEnabled((enabled) => !enabled)}
              >
                <span>גופן קריא</span>
                <span className="accessibility-option-state">
                  {readableFontEnabled ? "פעיל" : "כבוי"}
                </span>
              </button>
              <a
                className="accessibility-skip-link"
                href="#catalog"
                onClick={() => setAccessibilityOpen(false)}
              >
                מעבר ישיר לקטלוג
              </a>
              <button
                className="accessibility-reset"
                type="button"
                disabled={
                  !highContrastEnabled &&
                  !underlinedLinksEnabled &&
                  accessibilityTextScale === 1 &&
                  !strongFocusEnabled &&
                  !reducedMotionEnabled &&
                  !expandedSpacingEnabled &&
                  !readableFontEnabled
                }
                onClick={() => {
                  setHighContrastEnabled(false);
                  setUnderlinedLinksEnabled(false);
                  setAccessibilityTextScale(1);
                  setStrongFocusEnabled(false);
                  setReducedMotionEnabled(false);
                  setExpandedSpacingEnabled(false);
                  setReadableFontEnabled(false);
                }}
              >
                איפוס אפשרויות
              </button>
            </section>
          )}
          <button
            className="accessibility-toggle"
            type="button"
            aria-label={
              accessibilityOpen ? "סגירת תפריט נגישות" : "פתיחת תפריט נגישות"
            }
            aria-expanded={accessibilityOpen}
            aria-controls={
              accessibilityOpen ? "accessibility-panel" : undefined
            }
            onClick={() => setAccessibilityOpen((open) => !open)}
          >
            <svg
              viewBox="0 0 32 32"
              fill="none"
              aria-hidden="true"
              focusable="false"
            >
              <circle cx="16" cy="5.6" r="2.7" fill="currentColor" />
              <path
                d="M6 11.2h20M16 9v8m0 0-6.2 9M16 17l6.2 9M9.8 11.2l2.4 6.1m10-6.1-2.4 6.1"
                stroke="currentColor"
                strokeWidth="2.4"
                strokeLinecap="round"
                strokeLinejoin="round"
              />
            </svg>
          </button>
        </div>
      )}

      {selectedProduct && (
        <div
          className="product-details-overlay"
          onMouseDown={(event) => {
            if (event.target === event.currentTarget) {
              setSelectedProduct(null);
            }
          }}
        >
          <section
            className="product-details-dialog"
            role="dialog"
            aria-modal="true"
            aria-labelledby="product-details-title"
          >
            <header className="product-details-header">
              <div>
                <p className="eyebrow section-eyebrow">פרטי מוצר</p>
                <h2 id="product-details-title">
                  {selectedProduct.description}
                </h2>
              </div>
              <button
                className="cart-close"
                type="button"
                autoFocus
                aria-label="סגירת פרטי המוצר"
                onClick={() => setSelectedProduct(null)}
              >
                ×
              </button>
            </header>
            <div className="product-details-layout">
              <div className="product-detail-image-frame">
                {productImageBySku[selectedProduct.sku] ? (
                  <img
                    className="product-detail-image"
                    src={productImageBySku[selectedProduct.sku]}
                    alt={selectedProduct.description}
                  />
                ) : (
                  <div className="product-detail-placeholder">
                    <span aria-hidden="true">✳</span>
                    {categoryLabels[getCategory(selectedProduct.description)]}
                  </div>
                )}
              </div>
              <div className="product-detail-info">
                <dl className="product-detail-specs">
                  <div>
                    <dt>מק״ט</dt>
                    <dd>{selectedProduct.sku}</dd>
                  </div>
                  <div>
                    <dt>אריזה</dt>
                    <dd>{selectedProduct.capacity || "לא צוינה"}</dd>
                  </div>
                  <div>
                    <dt>ברקוד</dt>
                    <dd>{selectedProduct.barcode || "לא צוין"}</dd>
                  </div>
                  <div>
                    <dt>מחיר ליחידה</dt>
                    <dd>
                      {selectedProductPrice === null
                        ? "לא זמין"
                        : formatCurrency(selectedProductPrice)}
                    </dd>
                  </div>
                  <div>
                    <dt>כפולת הזמנה</dt>
                    <dd>
                      {getOrderMultiple(selectedProduct.capacity)} יחידות
                    </dd>
                  </div>
                </dl>

                {selectedProductVariants.length > 1 && (
                  <div className="related-products">
                    <h3>אפשרויות נוספות</h3>
                    <div className="related-product-list">
                      {selectedProductVariants.map((variant) => (
                        <button
                          className={`related-product-option${variant.sku === selectedProduct.sku ? " active" : ""}`}
                          key={variant.sku}
                          type="button"
                          aria-pressed={variant.sku === selectedProduct.sku}
                          onClick={() => setSelectedProduct(variant)}
                        >
                          <span>{variant.description}</span>
                          <small>מק״ט {variant.sku}</small>
                        </button>
                      ))}
                    </div>
                  </div>
                )}

                <button
                  className="product-detail-add"
                  type="button"
                  onClick={() => addToCart(selectedProduct)}
                >
                  <span>
                    {selectedProductQuantity > 0
                      ? `הוספת אריזה לעגלה · בעגלה ${selectedProductQuantity}`
                      : "הוספה לעגלה"}
                  </span>
                  <small>
                    כפולות של {getOrderMultiple(selectedProduct.capacity)}
                  </small>
                </button>
              </div>
            </div>
          </section>
        </div>
      )}

      {cartOpen && (
        <div
          className="cart-overlay"
          onMouseDown={(event) => {
            if (event.target === event.currentTarget) {
              setCartOpen(false);
            }
          }}
        >
          <section
            className="cart-drawer"
            role="dialog"
            aria-modal="true"
            aria-labelledby="cart-title"
          >
            <header className="cart-header">
              <div>
                <p className="eyebrow section-eyebrow">ההזמנה שלכם</p>
                <h2 id="cart-title">עגלת קניות</h2>
              </div>
              <button
                className="cart-close"
                type="button"
                onClick={() => setCartOpen(false)}
                aria-label="סגירת העגלה"
              >
                ×
              </button>
            </header>

            {cart.length > 0 ? (
              <>
                <div className="cart-items">
                  {cart.map(({ product, quantity }) => {
                    const multiple = getOrderMultiple(product.capacity);
                    const numericQuantity = Number(quantity);
                    const validQuantity =
                      Number.isInteger(numericQuantity) &&
                      numericQuantity >= multiple &&
                      numericQuantity % multiple === 0;
                    const unitPrice = getNumericPrice(product.price);
                    const lineTotal =
                      validQuantity && unitPrice !== null
                        ? unitPrice * numericQuantity
                        : null;

                    return (
                      <article
                        className="cart-item"
                        key={product.sku}
                      >
                        <div className="cart-item-info">
                          <span className="cart-item-sku">
                            מק״ט {product.sku}
                          </span>
                          <h3>{product.description}</h3>
                          <span className="cart-item-capacity">
                            אריזה: {product.capacity}
                          </span>
                        </div>
                        <div className="cart-item-bottom">
                          <div>
                            <div className="quantity-control">
                              <button
                                type="button"
                                aria-label={`הפחתת כמות ${product.description}`}
                                onClick={() =>
                                  adjustCartQuantity(product.sku, -1)
                                }
                              >
                                −
                              </button>
                              <input
                                aria-label={`כמות ${product.description}`}
                                type="number"
                                min={multiple}
                                step={multiple}
                                value={quantity}
                                aria-invalid={!validQuantity}
                                onChange={(event) =>
                                  updateCartQuantity(
                                    product.sku,
                                    event.target.value,
                                  )
                                }
                              />
                              <button
                                type="button"
                                aria-label={`הגדלת כמות ${product.description}`}
                                onClick={() =>
                                  adjustCartQuantity(product.sku, 1)
                                }
                              >
                                +
                              </button>
                            </div>
                            {!validQuantity && (
                              <p className="quantity-error" role="alert">
                                הכמות צריכה להיות בכפולות של {multiple}
                              </p>
                            )}
                          </div>
                          <div className="cart-item-price">
                            {lineTotal === null
                              ? unitPrice === null
                                ? "מחיר לא זמין"
                                : "—"
                              : formatCurrency(lineTotal)}
                          </div>
                          <button
                            className="remove-item"
                            type="button"
                            onClick={() => removeFromCart(product.sku)}
                            aria-label={`הסרת ${product.description} מהעגלה`}
                          >
                            הסרה
                          </button>
                        </div>
                        <p className="cart-item-unit-price">
                          {unitPrice === null
                            ? "מחיר לא זמין"
                            : `${formatCurrency(unitPrice)} ליחידה`}
                          <span className="print-quantity">
                            כמות: {quantity}
                          </span>
                        </p>
                      </article>
                    );
                  })}
                </div>

                <div className="cart-summary">
                  <div className="cart-total-row cart-subtotal-row">
                    <span>סה״כ לפני מע״מ</span>
                    <strong>{formatCurrency(cartSubtotal)}</strong>
                  </div>
                  <div className="cart-total-row cart-vat-row">
                    <span>מע״מ {VAT_RATE * 100}%</span>
                    <strong>{formatCurrency(cartVat)}</strong>
                  </div>
                  <div className="cart-total-row cart-grand-total-row">
                    <span>סה״כ כולל מע״מ</span>
                    <strong>{formatCurrency(cartTotalWithVat)}</strong>
                  </div>
                  {cart.some(({ product }) => getNumericPrice(product.price) === null) && (
                    <p className="cart-note">
                      חלק מהמוצרים ללא מחיר בקובץ, ולכן אינם כלולים בסכום.
                    </p>
                  )}
                  <div className="cart-export-actions">
                    <button
                      type="button"
                      className="export-button excel-button"
                      disabled={hasInvalidQuantity}
                      onClick={() =>
                        downloadCartExcel(cartRows, {
                          subtotal: cartSubtotal,
                          vat: cartVat,
                          totalWithVat: cartTotalWithVat,
                        })
                      }
                    >
                      הורדת Excel
                    </button>
                    <button
                      type="button"
                      className="export-button pdf-button"
                      disabled={hasInvalidQuantity}
                      onClick={printCartAsPdf}
                    >
                      שמירה כ־PDF
                    </button>
                  </div>
                  {hasInvalidQuantity && (
                    <p className="export-error" role="alert">
                      יש לתקן את הכמויות לפני ייצוא ההזמנה.
                    </p>
                  )}
                </div>
              </>
            ) : (
              <div className="cart-empty">
                <span aria-hidden="true">✳</span>
                <h3>העגלה עוד ריקה</h3>
                <p>הוסיפו מוצרים מהקטלוג כדי להתחיל הזמנה.</p>
                <button
                  type="button"
                  onClick={() => {
                    setCartOpen(false);
                    document
                      .getElementById("catalog")
                      ?.scrollIntoView({ behavior: "smooth" });
                  }}
                >
                  מעבר לקטלוג
                </button>
              </div>
            )}
          </section>
        </div>
      )}
    </main>
  );
}

export default App;
