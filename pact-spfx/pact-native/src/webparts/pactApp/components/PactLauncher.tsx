import * as React from 'react';
import styles from './PactLauncher.module.css';

// ─── Types ────────────────────────────────────────────────────────────────────

/** Descriptor for a single launcher tile. */
export interface IAppTileConfig {
  /** Unique key — used as React key and aria-label base. */
  id: string;
  /** Card heading displayed below the logo. */
  title: string;
  /** Optional short description displayed between title and button. */
  description?: string;
  /** Path to image asset (use require()). Pass undefined to render no logo. */
  logoSrc?: string;
  /** Alt text for the logo image. */
  logoAlt?: string;
  /** Button label. Defaults to "Open". */
  buttonLabel?: string;
  /** Invoked when the card or its button is clicked. */
  onClick: () => void;
}

/** Props for the AppTile sub-component. */
type AppTileProps = IAppTileConfig;

// ─── AppTile ─────────────────────────────────────────────────────────────────

/**
 * A single Microsoft 365-style app launcher tile.
 *
 * Keyboard accessible: the card itself is focusable and responds to Enter/Space.
 * The inner button also fires `onClick` independently (with propagation stopped
 * to prevent double-firing).
 */
const AppTile: React.FC<AppTileProps> = ({
  id,
  title,
  description,
  logoSrc,
  logoAlt,
  buttonLabel = 'Open',
  onClick,
}) => {
  const handleCardKeyDown = (e: React.KeyboardEvent<HTMLDivElement>): void => {
    if (e.key === 'Enter' || e.key === ' ') {
      e.preventDefault();
      onClick();
    }
  };

  const handleButtonClick = (e: React.MouseEvent<HTMLButtonElement>): void => {
    e.stopPropagation(); // prevent card's onClick from also firing
    onClick();
  };

  return (
    <div
      className={styles.card}
      role="button"
      tabIndex={0}
      aria-label={`Open ${title}`}
      onClick={onClick}
      onKeyDown={handleCardKeyDown}
    >
      {/* Logo */}
      {logoSrc && (
        <div className={styles.logoWrapper} aria-hidden="true">
          <img
            src={logoSrc}
            alt={logoAlt ?? title}
            className={styles.logo}
          />
        </div>
      )}

      {/* Title */}
      <h2 className={styles.title}>{title}</h2>

      {/* Optional description */}
      {description && (
        <p className={styles.description}>{description}</p>
      )}

      {/* Action button */}
      <button
        type="button"
        className={styles.actionButton}
        aria-label={`${buttonLabel} ${title}`}
        onClick={handleButtonClick}
      >
        {buttonLabel}
      </button>
    </div>
  );
};

// ─── PactLauncher ─────────────────────────────────────────────────────────────

/** Props accepted by the PactLauncher component. */
export interface IPactLauncherProps {
  /** Called when the user clicks a tile or its button. */
  onOpen: () => void;
}

/**
 * PactLauncher — the idle startup view of the PACT SharePoint web part.
 *
 * Renders a responsive CSS Grid of Microsoft 365-style app launcher tiles.
 * Currently shows one tile (PACT Portal). Additional tiles can be added by
 * extending the `tiles` array below without changing any markup or CSS.
 *
 * This component is deliberately stateless — all state lives in PactApp.
 */
export const PactLauncher: React.FC<IPactLauncherProps> = ({ onOpen }) => {
  /**
   * Tile definitions.
   * To add more launcher tiles in future, push additional IAppTileConfig
   * objects to this array. The grid adapts automatically.
   */
  const tiles: IAppTileConfig[] = [
    {
      id: 'pact-portal',
      title: 'PACT Portal',
      description: 'Manage cases, appeals, policies and staff records.',
      // eslint-disable-next-line @typescript-eslint/no-var-requires
      logoSrc: require('../pact/assets/kcc-logo.png'),
      logoAlt: 'Konstructum KCC logo',
      buttonLabel: 'Open',
      onClick: onOpen,
    },
    // Future tiles — uncomment and fill in to add more app tiles:
    // {
    //   id: 'product-portal',
    //   title: 'Product Portal',
    //   description: 'Browse and manage product catalogue.',
    //   logoSrc: require('../pact/assets/product-logo.png'),
    //   logoAlt: 'Product Portal logo',
    //   buttonLabel: 'Open',
    //   onClick: () => { /* navigate */ },
    // },
  ];

  return (
    <section className={styles.launcherSection} aria-label="PACT application launcher">
      <div className={styles.cardsContainer} role="list">
        {tiles.map((tile) => (
          <div key={tile.id} role="listitem">
            <AppTile {...tile} />
          </div>
        ))}
      </div>
    </section>
  );
};

export default PactLauncher;
