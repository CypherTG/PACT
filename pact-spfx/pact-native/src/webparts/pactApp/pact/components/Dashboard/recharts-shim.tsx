/**
 * Recharts v3 ships @types/react 18 signatures whose return type
 * (`ReactNode`) includes `undefined`, which React 17's JSX.Element
 * definition rejects.  This shim casts each component to
 * React.ComponentType<any> so the rest of the codebase compiles
 * cleanly under SPFx's React 17 toolchain.
 */
import type { ComponentType } from 'react';
import {
  AreaChart  as _AreaChart,
  Area       as _Area,
  XAxis      as _XAxis,
  YAxis      as _YAxis,
  CartesianGrid as _CartesianGrid,
  Tooltip    as _Tooltip,
  ResponsiveContainer as _ResponsiveContainer,
  BarChart   as _BarChart,
  Bar        as _Bar,
  Legend     as _Legend,
} from 'recharts';

// eslint-disable-next-line @typescript-eslint/no-explicit-any
const cast = <T,>(c: T): ComponentType<any> => c as any;

export const AreaChart           = cast(_AreaChart);
export const Area                = cast(_Area);
export const XAxis               = cast(_XAxis);
export const YAxis               = cast(_YAxis);
export const CartesianGrid       = cast(_CartesianGrid);
export const RechartsTooltip     = cast(_Tooltip);
export const ResponsiveContainer = cast(_ResponsiveContainer);
export const BarChart            = cast(_BarChart);
export const Bar                 = cast(_Bar);
export const Legend              = cast(_Legend);
