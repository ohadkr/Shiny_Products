import { useEffect, useMemo, useState } from "react";
import productsData from "./data/products.json";
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

type CategoryId =
  | "all"
  | "cloths"
  | "wipes"
  | "scrubbers"
  | "tools"
  | "bags"
  | "gloves"
  | "other";

const products: Product[] = productsData;
const PAGE_SIZE = 24;

const categoryLabels: Record<CategoryId, string> = {
  all: "כל המוצרים",
  cloths: "מטליות ומגבות",
  wipes: "מגבונים",
  scrubbers: "ספוגים וקרצוף",
  tools: "כלי ניקיון",
  bags: "שקיות",
  gloves: "כפפות",
  other: "מוצרים נוספים",
};

const categoryOrder: CategoryId[] = [
  "all",
  "cloths",
  "wipes",
  "scrubbers",
  "tools",
  "bags",
  "gloves",
  "other",
];

function getCategory(description: string): CategoryId {
  const name = description.toLowerCase();

  if (/מגבונים|מגבון|לח(ים|ות) לרצפה|נייר טואלט לח/.test(name)) {
    return "wipes";
  }
  if (/מטליות|מטלית|מגבת|סחבה|גליל מטליות/.test(name)) {
    return "cloths";
  }
  if (/כריות|כרית|ספוג|ספונג|קרצוף/.test(name)) {
    return "scrubbers";
  }
  if (/שקיות|שקית|אשפה/.test(name)) {
    return "bags";
  }
  if (/כפפות|כפפה/.test(name)) {
    return "gloves";
  }
  if (/מגב|מטאטא|דלי|מברשת|יעה|מקל|מתקן/.test(name)) {
    return "tools";
  }

  return "other";
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
}: {
  product: Product;
  index: number;
  quantityInCart: number;
  onAdd: (product: Product) => void;
  onAdjust: (sku: string, direction: -1 | 1) => void;
}) {
  const category = getCategory(product.description);
  const number = String(index + 1).padStart(2, "0");
  const orderMultiple = getOrderMultiple(product.capacity);

  return (
    <article className="product-card">
      <div className={`product-visual visual-${category}`}>
        <span className="visual-orbit orbit-one" />
        <span className="visual-orbit orbit-two" />
        <span className="visual-center" aria-hidden="true">
          <svg viewBox="0 0 48 48" fill="none">
            <path d="M24 5 28.7 18.3 42 23l-13.3 4.7L24 41l-4.7-13.3L6 23l13.3-4.7L24 5Z" />
          </svg>
        </span>
        <span className="visual-index">{number}</span>
        <span className="visual-category">{categoryLabels[category]}</span>
      </div>
      <div className="product-copy">
        <div className="product-topline">
          <span className="product-sku">מק״ט {product.sku || "—"}</span>
          {product.price && (
            <span className="product-price">
              <span className="meta-label">מחיר</span>
              {formatCurrency(Number(product.price))}
            </span>
          )}
        </div>
        <h2>{product.description || "מוצר שייני"}</h2>
        <div className="product-meta">
          {product.capacity && (
            <span>
              <span className="meta-label">אריזה</span>
              {product.capacity}
            </span>
          )}
          {product.barcode && (
            <span className="barcode">
              <span className="meta-label">ברקוד</span>
              {product.barcode}
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
  const [selectedCategory, setSelectedCategory] = useState<CategoryId>("all");
  const [visibleCount, setVisibleCount] = useState(PAGE_SIZE);
  const [cart, setCart] = useState<CartItem[]>([]);
  const [cartOpen, setCartOpen] = useState(false);

  useEffect(() => {
    if (!cartOpen) {
      return;
    }

    function closeOnEscape(event: KeyboardEvent) {
      if (event.key === "Escape") {
        setCartOpen(false);
      }
    }

    document.addEventListener("keydown", closeOnEscape);
    return () => document.removeEventListener("keydown", closeOnEscape);
  }, [cartOpen]);

  const categoryCounts = useMemo(() => {
    const counts = Object.fromEntries(
      categoryOrder.map((category) => [category, 0]),
    ) as Record<CategoryId, number>;

    products.forEach((product) => {
      counts[getCategory(product.description)] += 1;
    });
    counts.all = products.length;
    return counts;
  }, []);

  const filteredProducts = useMemo(() => {
    const normalizedQuery = query.trim().toLocaleLowerCase("he");

    return products.filter((product) => {
      const matchesCategory =
        selectedCategory === "all" ||
        getCategory(product.description) === selectedCategory;
      const searchableText = [
        product.description,
        product.sku,
        product.capacity,
        product.barcode,
      ]
        .join(" ")
        .toLocaleLowerCase("he");

      return matchesCategory && searchableText.includes(normalizedQuery);
    });
  }, [query, selectedCategory]);

  const visibleProducts = filteredProducts.slice(0, visibleCount);
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
  const cartTotal =
    Math.round(
      cartRows.reduce((total, row) => total + (row.lineTotal ?? 0), 0) * 100,
    ) / 100;
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
  }

  function updateCategory(category: CategoryId) {
    setSelectedCategory(category);
    setVisibleCount(PAGE_SIZE);
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

  return (
    <main>
      <header className="site-header">
        <a
          className="brand"
          href="#top"
          aria-label="שייני - חזרה לראש העמוד"
          id="top"
        >
          <img className="brand-logo" src="/shiny-logo.png" alt="Shiny" />
          <span className="brand-divider" />
          <span className="brand-caption">מוצרי ניקיון</span>
        </a>
        <div className="header-actions">
          <a className="header-link" href="#catalog">
            לקטלוג <span aria-hidden="true">↓</span>
          </a>
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
      </header>

      <section className="hero" aria-labelledby="hero-title">
        <div className="hero-copy">
          <p className="eyebrow">
            <span className="eyebrow-dot" />
            הבית הנקי מתחיל כאן
          </p>
          <h1 id="hero-title">
            כל מה שצריך
            <br />
            <span>לנקות בשמחה.</span>
          </h1>
          <p className="hero-description">
            מוצרים חכמים לניקיון יומיומי, לבית ולעסק.
            <br className="desktop-break" />
            בחרו את מה שמתאים לכם מתוך הקטלוג שלנו.
          </p>
          <a className="hero-action" href="#catalog">
            לגלות את הקטלוג <span aria-hidden="true">←</span>
          </a>
        </div>
        <div className="hero-art" aria-hidden="true">
          <div className="art-glow" />
          <div className="art-ring ring-back" />
          <div className="art-ring ring-front" />
          <div className="art-spark spark-one">✳</div>
          <div className="art-spark spark-two">✦</div>
          <div className="art-card art-card-main">
            <span className="art-card-small">SHINY · EVERYDAY CLEAN</span>
            <span className="art-card-flower">✳</span>
            <span className="art-card-name">ניקיון<br />עם חיוך</span>
          </div>
          <div className="art-card art-card-note">
            <span className="note-dot" />
            קטן בפרטים,
            <br />
            גדול בניקיון
          </div>
        </div>
        <div className="hero-footnote">
          <span className="footnote-line" />
          <span>נעים להכיר, אנחנו שייני</span>
        </div>
      </section>

      <section className="catalog-section" id="catalog">
        <div className="catalog-heading">
          <div>
            <p className="eyebrow section-eyebrow">הקטלוג שלנו</p>
            <h2>מוצאים את הניקיון שלכם</h2>
          </div>
          <div className="catalog-total">
            <span className="total-number">{products.length}</span>
            <span>מוצרים לבחור מהם</span>
          </div>
        </div>

        <div className="catalog-controls">
          <div className="search-wrap">
            <SearchIcon />
            <input
              aria-label="חיפוש מוצרים"
              type="search"
              value={query}
              onChange={(event) => updateQuery(event.target.value)}
              placeholder="מה מחפשים? שם מוצר, מק״ט או ברקוד"
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
          </div>
          <p className="result-count" aria-live="polite">
            {filteredProducts.length} מוצרים
          </p>
        </div>

        <div className="category-list" aria-label="סינון לפי קטגוריה">
          {categoryOrder
            .filter((category) => category === "all" || categoryCounts[category] > 0)
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

        {visibleProducts.length > 0 ? (
          <div className="product-grid">
            {visibleProducts.map((product, index) => (
              <ProductCard
                key={`${product.sku}-${product.barcode}-${index}`}
                product={product}
                index={filteredProducts.indexOf(product)}
                quantityInCart={
                  Number(
                    cart.find((item) => item.product.sku === product.sku)
                      ?.quantity,
                  ) || 0
                }
                onAdd={addToCart}
                onAdjust={adjustCartQuantity}
              />
            ))}
          </div>
        ) : (
          <div className="empty-state">
            <span className="empty-icon" aria-hidden="true">✳</span>
            <h3>לא מצאנו מוצרים מתאימים</h3>
            <p>נסו לחפש במילים אחרות או לבחור קטגוריה אחרת.</p>
            <button
              type="button"
              onClick={() => {
                updateQuery("");
                updateCategory("all");
              }}
            >
              הצגת כל המוצרים
            </button>
          </div>
        )}

        {visibleCount < filteredProducts.length && (
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

      <footer className="site-footer">
        <a className="footer-brand" href="#top" aria-label="שייני - חזרה לראש העמוד">
          <img src="/shiny-logo.png" alt="Shiny" />
        </a>
        <span>עושים מקום לניקיון טוב יותר.</span>
        <a href="#top">בחזרה למעלה ↑</a>
      </footer>

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
                  <div className="cart-total-row">
                    <span>סה״כ לפריטים עם מחיר</span>
                    <strong>{formatCurrency(cartTotal)}</strong>
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
                      onClick={() => downloadCartExcel(cartRows, cartTotal)}
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
