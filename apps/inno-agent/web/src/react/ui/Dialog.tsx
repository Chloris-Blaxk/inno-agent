import { useLayoutEffect, useRef, type ReactNode } from "react";
import { createPortal } from "react-dom";

/** Native modal focus containment, Escape handling and focus restoration. */
export function Dialog({ children, label, className, onClose }: {
	children: ReactNode;
	label: string;
	className?: string;
	onClose: () => void;
}) {
	const ref = useRef<HTMLDialogElement>(null);
	useLayoutEffect(() => {
		const dialog = ref.current;
		dialog?.showModal();
		return () => { if (dialog?.open) dialog.close(); };
	}, []);
	return createPortal(
		<dialog ref={ref} aria-label={label} className={className}
			onCancel={(event) => { event.preventDefault(); onClose(); }}
			onClick={(event) => { if (event.target === event.currentTarget) onClose(); }}>
			{children}
		</dialog>,
		document.body,
	);
}
