using System;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.InteropServices;

namespace SimplePPSX
{
	/// <summary>
	/// Late-bound (IDispatch) access to the presentation application's object model.
	///
	/// WHY THIS EXISTS
	/// ===============
	/// This project does not reference the real PowerPoint PIA. It carries its own
	/// hand-written partial declaration of it, under Microsoft\Office\Interop\PowerPoint:
	/// roughly 30 [ComImport] interfaces, each carrying the Microsoft interface GUID
	/// (Application is 91493442-5A91-11CF-8700-00AA0060263B, and so on).
	///
	/// Casting a COM object to a [ComImport] interface is a QueryInterface for that
	/// interface's IID. PowerPoint of course supports them. WPS Presentation does not:
	/// it publishes a VBA-compatible IDispatch surface with the same member NAMES, but
	/// its own CLSIDs and type library, so QueryInterface for a Microsoft IID fails and
	/// the cast throws InvalidCastException. That is why WPS support could not simply be
	/// switched on: nothing about the cast is negotiable at runtime.
	///
	/// IDispatch, by contrast, resolves members by name through GetIDsOfNames. WPS is
	/// documented as VBA-compatible and the member names are identical, so going
	/// through IDispatch works on both products. That is the whole reason this class
	/// exists, and the reason the ink-clearing code was already written this way (see
	/// PptOperator.IsInkShapeType, whose comment says exactly that).
	///
	/// WHY NOT JUST USE `dynamic`
	/// ==========================
	/// `dynamic` would be shorter, and is used where it reads best, but it cannot be
	/// used here for three reasons:
	///   - a missing member on a COM object surfaces as RuntimeBinderException from
	///     arbitrary call sites instead of one place, so a WPS gap would be very hard
	///     to locate later,
	///   - every call would need its own try/catch, and this object model throws on
	///     ordinary races (the show ends while a call is in flight),
	///   - the callers already read the results into typed locals, and this keeps that.
	/// The trade is a little verbosity for one choke point and one log.
	///
	/// ERROR HANDLING
	/// ==============
	/// Failures return a fallback rather than propagating, because this object model
	/// throws routinely for benign reasons: the slideshow ended, the presentation was
	/// closed, or WPS simply does not implement a member PowerPoint has. Every failure
	/// is logged once at Debug level so a real gap can still be diagnosed, and so a
	/// WPS-only failure does not look identical to "the show ended".
	///
	/// All members are static and take the target as the first argument, which keeps
	/// them usable with the `dynamic` values the rest of the code already holds.
	/// </summary>
	internal static class ComLate
	{
		/// <summary>Where a failure is recorded. Set once at startup to the app logger.</summary>
		internal static Action<string> LogFailure { get; set; }

		// MsoTriState: msoFalse = 0. Spelled out rather than referenced, because the
		// enum comes from the hand-written Microsoft.Office.Core declaration and this
		// class deliberately has no dependency on any typed interop at all.
		private const int MsoFalse = 0;

		/// <summary>Reads a property. Returns fallback when it cannot be read.</summary>
		internal static object Get(object target, string member, object fallback = null)
		{
			if (target == null)
			{
				return fallback;
			}
			try
			{
				// Type.InvokeMember with GetProperty, NOT `((dynamic)target)[member]`.
				//
				// The indexer form is the trap. The C# binder resolves d["Name"] as an
				// indexed property access and issues DISPATCH_PROPERTYGET with a named
				// DISPPARAM, which PowerPoint rejects outright:
				//
				//   d.Path          -> works
				//   d["Path"]       -> COMException 0x80020011 DISP_E_NOTACOLLECTION
				//
				// Verified against a live PowerPoint 16.0. Since this helper is the single
				// point through which the whole object model is reached, using the wrong
				// form here silently broke EVERY member access - which is exactly what
				// happened: the toolbar never appeared because SlideShowWindows came back
				// null, and the log filled with DISP_E_NOTACOLLECTION.
				//
				// InvokeMember goes through IDispatch::GetIDsOfNames and dispatches as a
				// plain property get, which is what both PowerPoint and WPS expect.
				return target.GetType().InvokeMember(member, BindingFlags.GetProperty, null, target, null);
			}
			catch (Exception ex)
			{
				Report("Get " + member, ex);
				return fallback;
			}
		}

		/// <summary>
		/// Reads a property and does NOT report a failure.
		///
		/// For members whose absence is an expected outcome rather than a fault, so
		/// probing for one cannot fill the log.
		/// </summary>
		internal static object GetQuiet(object target, string member)
		{
			if (target == null)
			{
				return null;
			}
			try
			{
				return target.GetType().InvokeMember(member, BindingFlags.GetProperty, null, target, null);
			}
			catch (Exception)
			{
				return null;
			}
		}

		/// <summary>
		/// Reads a property as an int and does NOT report a failure.
		///
		/// For the handful of reads whose absence is expected and harmless, where a log
		/// line would be noise: Application.HWND (used only to raise the app's own
		/// window) and SameObject's identity probes.
		/// </summary>
		internal static int GetIntQuiet(object target, string member, int fallback = 0)
		{
			if (target == null)
			{
				return fallback;
			}
			try
			{
				object value = target.GetType().InvokeMember(member, BindingFlags.GetProperty, null, target, null);
				return value == null ? fallback : Convert.ToInt32(value);
			}
			catch (Exception)
			{
				return fallback;
			}
		}

		/// <summary>Reads a property as an int. Returns fallback when it cannot be read.</summary>
		internal static int GetInt(object target, string member, int fallback = 0)
		{
			object value = ComLate.Get(target, member);
			if (value == null)
			{
				return fallback;
			}
			try
			{
				return Convert.ToInt32(value);
			}
			catch (Exception ex)
			{
				Report("GetInt " + member, ex);
				return fallback;
			}
		}

		/// <summary>Reads a property as a string. Returns fallback when it cannot be read.</summary>
		internal static string GetString(object target, string member, string fallback = null)
		{
			object value = ComLate.Get(target, member);
			return value == null ? fallback : value.ToString();
		}

		/// <summary>Writes a property. Reports and swallows a failure.</summary>
		internal static void Set(object target, string member, object value)
		{
			if (target == null)
			{
				return;
			}
			try
			{
				// InvokeMember rather than the dynamic indexer, for the reason documented
				// on Get: d["Name"] dispatches as an indexed access and PowerPoint refuses
				// it with DISP_E_NOTACOLLECTION. SetProperty dispatches as a plain put.
				target.GetType().InvokeMember(member, BindingFlags.SetProperty, null, target, new object[] { value });
			}
			catch (Exception ex)
			{
				Report("Set " + member, ex);
			}
		}

		/// <summary>Calls a method. Reports and swallows a failure.</summary>
		internal static void Call(object target, string method, params object[] args)
		{
			if (target == null)
			{
				return;
			}
			try
			{
				target.GetType().InvokeMember(method, BindingFlags.InvokeMethod, null, target, args);
			}
			catch (Exception ex)
			{
				Report("Call " + method, ex);
			}
		}

		/// <summary>
		/// GotoSlide, spelled out because its second argument is optional in the object
		/// model and passing it through params would send an empty array rather than
		/// "omit it". SlideShowView.GotoSlide(Index, ResetSlide) needs msoFalse to stop
		/// animations restarting, which is the whole reason this app jumps this way.
		/// </summary>
		internal static void CallGotoSlide(object view, int index)
		{
			if (view == null)
			{
				return;
			}
			try
			{
				((dynamic)view).GotoSlide(index, MsoFalse);
			}
			catch (Exception ex)
			{
				Report("Call GotoSlide", ex);
			}
		}

		/// <summary>
		/// Reads a collection's Count. Used instead of casting to IEnumerable because a
		/// COM collection reports Count through IDispatch but may not enumerate, and
		/// both products differ on which they support.
		/// </summary>
		internal static int CountOf(object collection)
		{
			return ComLate.GetInt(collection, "Count");
		}

		/// <summary>
		/// Reads an indexed item. PowerPoint collections are 1-based over IDispatch.
		/// </summary>
		internal static object Item(object collection, int index)
		{
			if (collection == null)
			{
				return null;
			}
			try
			{
				// An indexed access through the dynamic binder, and this is the ONE place
				// where the dynamic form is correct rather than a trap (see Get).
				// Indexed access is exactly what a COM collection's default member
				// wants: DISPATCH_PROPERTYGET with a positional argument.
				//
				// Verified against a live PowerPoint 16.0 SlideShowWindows collection:
				//   dynamic d[1]               -> works
				//   dynamic d.Item(1)          -> works
				//   InvokeMember("Item", args)  -> 0x80020003 (no such member)
				//   InvokeMember(null,   args)  -> ArgumentNullException
				// PowerPoint exposes no member literally named "Item", so the
				// InvokeMember spelling cannot work here at all.
				dynamic late = collection;
				return late[index];
			}
			catch (Exception ex)
			{
				Report("Item [" + index + "]", ex);
				return null;
			}
		}

		/// <summary>
		/// Builds a flat list out of a 1-based IDispatch collection.
		///
		/// Returns an empty list rather than null, so callers never need a null check
		/// before iterating, and so "the collection was unreadable" and "the
		/// presentation has no slides" behave the same way downstream.
		/// </summary>
		internal static List<object> ToList(object collection)
		{
			List<object> items = new List<object>();
			int count = CountOf(collection);
			for (int i = 1; i <= count; i++)
			{
				object item = Item(collection, i);
				if (item != null)
				{
					items.Add(item);
				}
			}
			return items;
		}

		/// <summary>
		/// True when two late-bound objects are the same COM object.
		///
		/// Reference equality is wrong here: IDispatch hands back a fresh RCW per
		/// call, so the same underlying COM object arrives as two different managed
		/// objects. Both products do expose the identity members, so this compares
		/// those and only falls back to reference equality if neither is readable.
		/// </summary>
		internal static bool SameObject(object left, object right)
		{
			if (ReferenceEquals(left, right))
			{
				return true;
			}
			if (left == null || right == null)
			{
				return false;
			}

			// Quiet: not every object in this model has an identity member (a Slide does,
			// a Shape may not), so a miss here is an expected outcome, not a fault.
			object leftId = ComLate.GetQuiet(left, "FullName");
			object rightId = ComLate.GetQuiet(right, "FullName");
			if (leftId != null && rightId != null)
			{
				return leftId.Equals(rightId);
			}

			leftId = ComLate.GetQuiet(left, "Name");
			rightId = ComLate.GetQuiet(right, "Name");
			return leftId != null && leftId.Equals(rightId);
		}

		private static void Report(string what, Exception ex)
		{
			Action<string> log = ComLate.LogFailure;
			if (log != null)
			{
				// Unwrap the binder's wrapper: the message that matters is COM's own
				// ("The parameter is incorrect", "Member not found"), not the
				// RuntimeBinderException text that hides it.
				Exception inner = ex;
				while (inner.InnerException != null && string.IsNullOrEmpty(inner.Message))
				{
					inner = inner.InnerException;
				}
				log("ComLate." + what + ": " + (inner.Message ?? ex.GetType().Name));
			}
		}
	}
}