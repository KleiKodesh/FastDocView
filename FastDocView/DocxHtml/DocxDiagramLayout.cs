using System.Globalization;
using System.Xml.Linq;

namespace FastDocView.DocxHtml;

/// <summary>
/// SmartArt laid out from its data model and layout definition, for a diagram saved without its drawing part (Word 16 always
/// writes one; other producers and stripped files do not). The result is a namespace-free dsp:spTree in the frame's EMU
/// space, drawn by DocxShapeReader like a saved drawing. A port of LibreOffice's diagram layout (oox/source/drawingml/diagram,
/// libreoffice-24-8): the layout tree (layoutnodecontext.cxx: layoutNode, alg + param, shape + adjLst, choose / if / else,
/// forEach + ref, constrLst, ruleLst), the shape-creation and layouting visitors (layoutatomvisitors.cxx: shapes first, then
/// per layout node constraints, rules, its algorithm, then the nested nodes), and the algorithms of AlgAtom::layoutShape
/// (diagramlayoutatoms.cxx: composite, conn, cycle, hierRoot / hierChild, sp, tx; lin, snake and pyra, which LibreOffice
/// places far from Word, follow ECMA-376 Part 1 21.4.2 and Word's own drawings instead). The loops walk the data
/// model along their axis as LibreOffice master does (diagramlayoutatoms.cxx stepOfAxis / childrenOf / parentOf /
/// isPointOfKind / pointsAlongAxis; layoutatomvisitorbase.cxx visit(ForEachAtom), presentationForPass): each pass stands on
/// a data point, and a layout node draws the presentation point built for it. Colours and line / fill / font references come
/// from the colors and quickStyle parts by the presentation point's style label (LayoutNode::setupShape), the n-th colour of a
/// list for the n-th point of that label (presStyleIdx). Text is the data points' text (presOf), fitted to the shape the way
/// the tx algorithm's autofit does: one size for every shape of a layout node, from its primFontSz constraint down to its rule.
/// </summary>
internal static class DocxDiagramLayout
{
    private const double EmuPerPoint = 12700;
    private const double EmuPerMillimetre = 36000;
    private static readonly char[] Space = [' '];

    /// <summary>The diagram's shapes as a namespace-free spTree in a frame of the given EMU size; null when the parts don't lay out.</summary>
    public static XElement? Lay(XElement data, XElement layoutDefinition, XElement? colors, XElement? quickStyle, double widthEmu, double heightEmu)
    {
        if (widthEmu <= 0 || heightEmu <= 0) return null;
        var model = new DataModel(data);
        if (model.Root is null) return null;
        if (layoutDefinition.Element("layoutNode") is not { } rootElement) return null;
        var layout = new LayoutTree();
        LayoutNodeAtom root = layout.ReadLayoutNode(rootElement, null, "b");
        var diagram = new Diagram(model, layout, colors, quickStyle);
        var frame = new DiagramShape { Name = root.Name, IsGroup = true, Width = widthEmu, Height = heightEmu };
        new ShapeCreationVisitor(diagram, model.Root, frame, root).Visit(root);
        new ShapeLayoutingVisitor(diagram, model.Root).Visit(root);
        SortByZOrder(frame);
        TextFitting.FitAll(frame);
        var tree = new XElement("spTree");
        Emit(frame, 0, 0, tree, diagram);
        return tree.HasElements ? tree : null;
    }

    // ── the data model (dgm:dataModel: dgm:ptLst, dgm:cxnLst) ─────────────────────────────────────────────────────────────

    internal sealed class Point
    {
        public string Id = "", Type = "node", PresName = "", PresAssocId = "", PresStyleLabel = "";
        public int PresStyleIndex = -1, PresStyleCount;
        public string Direction = "norm";
        public string? HierarchyBranch;
        public XElement? ShapeProperties, Text;
    }

    internal sealed class Connection
    {
        public string Id = "", Type = "parOf", Source = "", Destination = "", ParentTransition = "", SiblingTransition = "";
        public int SourceOrder, DestinationOrder;
    }

    internal sealed class DataModel
    {
        public readonly List<Point> Points = [];
        public readonly Dictionary<string, Point> ById = [];
        public readonly List<Connection> Connections = [];
        public readonly Point? Root;

        public DataModel(XElement data)
        {
            foreach (XElement element in data.Element("ptLst")?.Elements("pt") ?? [])
            {
                XElement? set = element.Element("prSet");
                var point = new Point
                {
                    Id = (string?)element.Attribute("modelId") ?? "", Type = (string?)element.Attribute("type") ?? "node",
                    PresName = (string?)set?.Attribute("presName") ?? "", PresAssocId = (string?)set?.Attribute("presAssocID") ?? "",
                    PresStyleLabel = (string?)set?.Attribute("presStyleLbl") ?? "",
                    PresStyleIndex = Int(set?.Attribute("presStyleIdx"), -1), PresStyleCount = Int(set?.Attribute("presStyleCnt"), 0),
                    ShapeProperties = element.Element("spPr"), Text = element.Element("t"),
                };
                XElement? variables = set?.Element("presLayoutVars");
                if ((string?)variables?.Element("dir")?.Attribute("val") is { Length: > 0 } direction) point.Direction = direction;
                point.HierarchyBranch = (string?)variables?.Element("hierBranch")?.Attribute("val");
                Points.Add(point);
                if (!ById.ContainsKey(point.Id)) ById[point.Id] = point;
            }
            foreach (XElement element in data.Element("cxnLst")?.Elements("cxn") ?? [])
                Connections.Add(new Connection
                {
                    Id = (string?)element.Attribute("modelId") ?? "", Type = (string?)element.Attribute("type") ?? "parOf",
                    Source = (string?)element.Attribute("srcId") ?? "", Destination = (string?)element.Attribute("destId") ?? "",
                    SourceOrder = Int(element.Attribute("srcOrd"), 0), DestinationOrder = Int(element.Attribute("destOrd"), 0),
                    ParentTransition = (string?)element.Attribute("parTransId") ?? "", SiblingTransition = (string?)element.Attribute("sibTransId") ?? "",
                });
            Root = Points.FirstOrDefault(p => p.Type == "doc");
        }

        public bool IsConnected(string from, string to) => Connections.Any(c => c.Source == from && c.Destination == to);

        /// <summary>childrenOf: what a parOf from <paramref name="from"/> holds of the wanted kind (the child, its parTrans or its sibTrans), in source order.</summary>
        public List<string> ChildrenOf(string from, string wanted)
            => Connections.Where(c => c.Type == "parOf" && c.Source == from).OrderBy(c => c.SourceOrder)
                          .Select(c => wanted == "parTrans" ? c.ParentTransition : wanted == "sibTrans" ? c.SiblingTransition : c.Destination)
                          .Where(id => id.Length > 0).ToList();

        /// <summary>parentOf: the data parent of a node or of one of its transitions, with the child's source order.</summary>
        public string? ParentOf(string from, out int order)
        {
            foreach (Connection connection in Connections)
                if (connection.Type == "parOf" && (connection.Destination == from || connection.SiblingTransition == from || connection.ParentTransition == from))
                {
                    order = connection.SourceOrder;
                    return connection.Source;
                }
            order = 0;
            return null;
        }

        public string TypeOf(string id) => ById.TryGetValue(id, out Point? point) ? point.Type : "";

        /// <summary>isPointOfKind: node covers assistants too, norm the ordinary node alone.</summary>
        public static bool IsOfKind(string type, string wanted) => wanted switch
        {
            "all" or "none" or "" => true,
            "node" => type is "node" or "asst",
            "norm" => type == "node",
            "asst" => type == "asst",
            "nonAsst" => type != "asst",
            "nonNorm" => type != "node",
            _ => type == wanted,
        };

        /// <summary>stepOfAxis: the points one step of an axis reaches.</summary>
        public List<string> Step(string from, string axis, string wanted)
        {
            switch (axis)
            {
                case "self": return [from];
                case "root": return Root is null ? [] : [Root.Id];
                case "ch": return ChildrenOf(from, wanted);
                case "par": return ParentOf(from, out _) is { } up ? [up] : [];
                case "followSib" or "follow" or "precedSib" or "preced":
                {
                    if (ParentOf(from, out int order) is not { } parent) return [];
                    bool transition = wanted is "parTrans" or "sibTrans";
                    bool following = axis is "followSib" or "follow";
                    return Connections.Where(c => c.Type == "parOf" && c.Source == parent)
                                      .Where(c => following ? (transition ? c.SourceOrder >= order : c.SourceOrder > order)
                                                            : (transition ? c.SourceOrder <= order : c.SourceOrder < order))
                                      .OrderBy(c => c.SourceOrder)
                                      .Select(c => wanted == "parTrans" ? c.ParentTransition : wanted == "sibTrans" ? c.SiblingTransition : c.Destination)
                                      .Where(id => id.Length > 0).ToList();
                }
                case "des" or "desOrSelf":
                {
                    var found = new List<string>();
                    if (axis == "desOrSelf") found.Add(from);
                    var seen = new HashSet<string>();
                    var stack = new Stack<string>();
                    stack.Push(from);
                    while (stack.Count > 0)
                    {
                        string current = stack.Pop();
                        if (!seen.Add(current)) continue;
                        List<string> children = ChildrenOf(current, "all");
                        found.AddRange(children);
                        for (int i = children.Count - 1; i >= 0; i--) stack.Push(children[i]);
                    }
                    return found;
                }
                default: return [];
            }
        }

        /// <summary>pointsAlongAxis: the axis walked one step at a time, each step filtered by its point type, start and count.</summary>
        public List<string> Along(Iterator iterator, string from, bool lastStepWhole)
        {
            var standing = new List<string> { from };
            for (int step = 0; step < iterator.Axes.Count; step++)
            {
                string wanted = step < iterator.PointTypes.Count ? iterator.PointTypes[step] : "all";
                var reached = new List<string>();
                foreach (string point in standing) reached.AddRange(Step(point, iterator.Axes[step], wanted));
                reached.RemoveAll(id => !IsOfKind(TypeOf(id), wanted));
                bool whole = lastStepWhole && step + 1 == iterator.Axes.Count;
                int start = !whole && step < iterator.Starts.Count ? iterator.Starts[step] : 0;
                int take = !whole && step < iterator.Counts.Count ? iterator.Counts[step] : 0;
                if (start > 1) reached.RemoveRange(0, Math.Min(start - 1, reached.Count));
                if (take > 0 && take < reached.Count) reached.RemoveRange(take, reached.Count - take);
                standing = reached;
            }
            return standing;
        }

        /// <summary>calcMaxDepth: the longest parOf path down from a point.</summary>
        public int MaxDepth(string id) => MaxDepth(id, []);

        private int MaxDepth(string id, HashSet<string> path)
        {
            if (!path.Add(id)) return 0;     // a parOf cycle (malformed data) ends the path
            int depth = Connections.Where(c => c.Type == "parOf" && c.Source == id).Select(c => MaxDepth(c.Destination, path) + 1).DefaultIfEmpty(0).Max();
            path.Remove(id);
            return depth;
        }

        /// <summary>calcDepth (svx datamodel.cxx): the parOf depth of a data point below the document.</summary>
        public int Depth(string id)
        {
            int depth = 0;
            for (int guard = 0; guard < 64; guard++)
            {
                Connection? up = Connections.FirstOrDefault(c => c.Type == "parOf" && c.Destination == id && c.ParentTransition.Length > 0 && c.SiblingTransition.Length > 0);
                if (up is null) break;
                depth++;
                id = up.Source;
            }
            return depth;
        }
    }

    // ── the layout definition (dgm:layoutDef) ─────────────────────────────────────────────────────────────────────────────

    internal abstract class Atom
    {
        public Atom? Parent;
        public readonly List<Atom> Children = [];
        public LayoutNodeAtom? OwnerNode => this as LayoutNodeAtom ?? Parent?.OwnerNode;
    }

    internal sealed class LayoutNodeAtom : Atom { public string Name = "", StyleLabel = "", ChildOrder = "b"; }
    internal sealed class ShapeAtom : Atom { public string? Type; public bool HideGeometry; public double Rotation; public int ZOrderOffset; public XElement? Adjustments; }
    internal sealed class ChooseAtom : Atom;
    internal sealed class ConstraintAtom : Atom { public Constraint Constraint = new(); }
    internal sealed class RuleAtom : Atom { public string ForName = "", Type = ""; public double Value; }

    internal sealed class AlgorithmAtom : Atom
    {
        public string Type = "";
        public readonly Dictionary<string, string> Parameters = [];
        public double AspectRatio;
        public string Parameter(string name, string fallback) => Parameters.TryGetValue(name, out string? value) ? value : fallback;
        public double Number(string name, double fallback)
            => Parameters.TryGetValue(name, out string? value) && double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double number) ? number : fallback;
    }

    internal sealed class Iterator
    {
        public List<string> Axes = [], PointTypes = [];
        public List<int> Starts = [], Counts = [];
        public int Count = -1, Start, Step = 1;
        public bool HideLastTransition = true;
        public string PointType => PointTypes.Count > 0 ? PointTypes[0] : "all";

        public static Iterator Read(XElement element)
        {
            List<int> Numbers(string name) => ((string?)element.Attribute(name) ?? "").Split(Space, StringSplitOptions.RemoveEmptyEntries)
                                                                                     .Select(v => int.TryParse(v, out int n) ? n : 0).ToList();
            var iterator = new Iterator
            {
                Axes = ((string?)element.Attribute("axis") ?? "").Split(Space, StringSplitOptions.RemoveEmptyEntries).ToList(),
                PointTypes = ((string?)element.Attribute("ptType") ?? "").Split(Space, StringSplitOptions.RemoveEmptyEntries).ToList(),
                Starts = Numbers("st"), Counts = Numbers("cnt"),
                HideLastTransition = (string?)element.Attribute("hideLastTrans") is not ("0" or "false"),
            };
            if (iterator.Counts.Count > 0) iterator.Count = iterator.Counts[0] == 0 ? -1 : iterator.Counts[0];
            if (iterator.Starts.Count > 0) iterator.Start = iterator.Starts[0];
            List<int> steps = Numbers("step");
            if (steps.Count > 0) iterator.Step = steps[0];
            return iterator;
        }
    }

    internal sealed class ForEachAtom : Atom { public string Name = "", Reference = ""; public Iterator Iterator = new(); }

    internal sealed class ConditionAtom : Atom
    {
        public bool IsElse;
        public Iterator Iterator = new();
        public string Function = "", Argument = "none", Operator = "", Value = "";
    }

    internal sealed class Constraint
    {
        public string For = "none", ForName = "", PointType = "none", Type = "none", RefFor = "none", RefForName = "", RefType = "none", RefPtType = "none", Operator = "none";
        public string? RefPointType() => RefPtType is "none" or "" ? null : RefPtType;
        public double Factor = 1, Value;
        public Constraint Copy() => (Constraint)MemberwiseClone();
    }

    internal sealed class LayoutTree
    {
        public readonly Dictionary<string, ForEachAtom> ForEachByName = [];

        public LayoutNodeAtom ReadLayoutNode(XElement element, Atom? parent, string inheritedChildOrder)
        {
            var node = new LayoutNodeAtom
            {
                Parent = parent, Name = (string?)element.Attribute("name") ?? "", StyleLabel = (string?)element.Attribute("styleLbl") ?? "",
                ChildOrder = (string?)element.Attribute("chOrder") ?? inheritedChildOrder,
            };
            ReadChildren(element, node, node);
            return node;
        }

        /// <summary>LayoutNodeContext::onCreateContext: the atoms of a layout node, forEach, if or else.</summary>
        private void ReadChildren(XElement element, Atom atom, LayoutNodeAtom owner)
        {
            foreach (XElement child in element.Elements())
            {
                switch (child.Name.LocalName)
                {
                    case "layoutNode": atom.Children.Add(ReadLayoutNode(child, atom, owner.ChildOrder)); break;
                    case "shape":
                        atom.Children.Add(new ShapeAtom
                        {
                            Parent = atom, Type = (string?)child.Attribute("type"), HideGeometry = IsTrue(child.Attribute("hideGeom")),
                            Rotation = Double(child.Attribute("rot"), 0), ZOrderOffset = Int(child.Attribute("zOrderOff"), 0), Adjustments = child.Element("adjLst"),
                        });
                        break;
                    case "alg":
                    {
                        var algorithm = new AlgorithmAtom { Parent = atom, Type = (string?)child.Attribute("type") ?? "" };
                        foreach (XElement parameter in child.Elements("param"))
                        {
                            string type = (string?)parameter.Attribute("type") ?? "";
                            if (type == "ar") algorithm.AspectRatio = Double(parameter.Attribute("val"), 0);
                            else algorithm.Parameters[type] = (string?)parameter.Attribute("val") ?? "";
                        }
                        atom.Children.Add(algorithm);
                        break;
                    }
                    case "choose":
                    {
                        var choose = new ChooseAtom { Parent = atom };
                        foreach (XElement branch in child.Elements().Where(e => e.Name.LocalName is "if" or "else"))
                        {
                            var condition = new ConditionAtom
                            {
                                Parent = choose, IsElse = branch.Name.LocalName == "else", Iterator = Iterator.Read(branch),
                                Function = (string?)branch.Attribute("func") ?? "", Argument = (string?)branch.Attribute("arg") ?? "none",
                                Operator = (string?)branch.Attribute("op") ?? "", Value = (string?)branch.Attribute("val") ?? "",
                            };
                            ReadChildren(branch, condition, owner);
                            choose.Children.Add(condition);
                        }
                        atom.Children.Add(choose);
                        break;
                    }
                    case "forEach":
                    {
                        var forEach = new ForEachAtom { Parent = atom, Name = (string?)child.Attribute("name") ?? "", Reference = (string?)child.Attribute("ref") ?? "", Iterator = Iterator.Read(child) };
                        if (forEach.Name.Length > 0) ForEachByName[forEach.Name] = forEach;
                        ReadChildren(child, forEach, owner);
                        atom.Children.Add(forEach);
                        break;
                    }
                    case "constrLst":
                        foreach (XElement constraint in child.Elements("constr"))
                            atom.Children.Add(new ConstraintAtom
                            {
                                Parent = atom,
                                Constraint = new Constraint
                                {
                                    For = (string?)constraint.Attribute("for") ?? "none", ForName = (string?)constraint.Attribute("forName") ?? "",
                                    PointType = (string?)constraint.Attribute("ptType") ?? "none", Type = (string?)constraint.Attribute("type") ?? "none",
                                    RefFor = (string?)constraint.Attribute("refFor") ?? "none", RefForName = (string?)constraint.Attribute("refForName") ?? "",
                                    RefType = (string?)constraint.Attribute("refType") ?? "none", RefPtType = (string?)constraint.Attribute("refPtType") ?? "none", Operator = (string?)constraint.Attribute("op") ?? "none",
                                    Factor = Double(constraint.Attribute("fact"), 1), Value = Double(constraint.Attribute("val"), 0),
                                },
                            });
                        break;
                    case "ruleLst":
                        foreach (XElement rule in child.Elements("rule"))
                            atom.Children.Add(new RuleAtom { Parent = atom, ForName = (string?)rule.Attribute("forName") ?? "", Type = (string?)rule.Attribute("type") ?? "", Value = Double(rule.Attribute("val"), 0) });
                        break;
                }
            }
        }
    }

    // ── shapes ────────────────────────────────────────────────────────────────────────────────────────────────────────────

    internal sealed class DiagramShape
    {
        public string Name = "";
        public bool IsGroup;
        public string? Preset;                 // a preset geometry; "conn" until the conn algorithm picks an arrow
        public bool HideGeometry;
        public XElement? Adjustments;
        public double X, Y, Width, Height, Rotation, AspectRatio;
        public int VerticalShapesCount = 1, ZOrderOffset, ZOrder;
        public string DataNodeType = "";
        public readonly List<DiagramShape> Children = [];
        public Point? Presentation;
        public string StyleLabel = "";
        public int ColorIndex;
        public readonly List<(XElement Paragraph, int Level)> Paragraphs = [];
        public readonly List<XElement> OwnProperties = [];
        public bool HasText => Paragraphs.Count > 0;
        // text, set by the tx algorithm
        public bool CentreText = true, BulletList;
        public int BulletStartLevel = 1;
        public string Anchor = "ctr";
        public double FontSize = 18, MinimumFontSize = 5;
        public (string Type, double Factor)[] Margins = [];
        public bool Autofit;
        public string AlgorithmType = "";      // the algorithm its layout node chose (hierRoot / hierChild rows of a tree)
        public bool HierarchyPlaced;           // laid out with the whole tree by the outermost hierarchy algorithm
        public bool FillsParent;               // the pyramid's level shape: its composite group's whole box
        public double BeginPad, EndPad;        // a connector's padding along its space, as fractions of it (begPad / endPad over connDist)
        public List<Constraint> OwnConstraints = [];
    }

    private sealed class Diagram(DataModel model, LayoutTree layout, XElement? colors, XElement? quickStyle)
    {
        public readonly DataModel Model = model;
        public readonly LayoutTree Layout = layout;
        public readonly XElement? Colors = colors, QuickStyle = quickStyle;
        public readonly Dictionary<Point, DiagramShape> ShapeOfPresentation = [];
        public readonly List<RuleAtom> Rules = [];
    }

    // ── the visitors (layoutatomvisitorbase.cxx / layoutatomvisitors.cxx) ─────────────────────────────────────────────────

    private abstract class Visitor(Diagram diagram, Point current)
    {
        protected readonly Diagram Diagram = diagram;
        protected Point Current = current;          // the presentation point reached (the document point at the start)
        protected string PassNodeId = "";           // the data point the pass of the enclosing loop stands on
        protected int CurrentIndex, CurrentStep, CurrentCount;
        private readonly HashSet<ForEachAtom> _followedReferences = [];

        public void Visit(Atom atom)
        {
            switch (atom)
            {
                case LayoutNodeAtom node: VisitLayoutNode(node); break;
                case ForEachAtom forEach: VisitForEach(forEach); break;
                case ChooseAtom choose: VisitChoose(choose); break;
                case ConditionAtom condition: DefaultVisit(condition); break;
                case AlgorithmAtom algorithm: VisitAlgorithm(algorithm); break;
                case ShapeAtom shape: VisitShape(shape); break;
                case ConstraintAtom constraint: VisitConstraint(constraint); break;
                case RuleAtom rule: VisitRule(rule); break;
            }
        }

        protected void DefaultVisit(Atom atom) { foreach (Atom child in atom.Children) Visit(child); }
        protected virtual void VisitLayoutNode(LayoutNodeAtom node) { }
        protected virtual void VisitAlgorithm(AlgorithmAtom algorithm) { }
        protected virtual void VisitShape(ShapeAtom shape) { }
        protected virtual void VisitConstraint(ConstraintAtom constraint) { }
        protected virtual void VisitRule(RuleAtom rule) { }

        private void VisitChoose(ChooseAtom choose)
        {
            foreach (ConditionAtom condition in choose.Children.OfType<ConditionAtom>())
                if (Decide(condition))
                {
                    Visit(condition);
                    break;
                }
        }

        private string WalkFrom => PassNodeId.Length > 0 ? PassNodeId : Current.Type == "pres" ? Current.PresAssocId : Current.Id;

        /// <summary>ConditionAtom::getDecision: var (dir, hierBranch), cnt, maxDepth, depth, pos / revPos / posEven / posOdd.</summary>
        private bool Decide(ConditionAtom condition)
        {
            if (condition.IsElse) return true;
            DataModel model = Diagram.Model;
            switch (condition.Function)
            {
                case "var":
                {
                    string? value = condition.Argument switch
                    {
                        "dir" => Current.Direction,
                        "hierBranch" => Current.HierarchyBranch ?? HierarchyBranchAbove(Current) ?? "std",
                        _ => null,
                    };
                    if (value is null) return true;
                    return condition.Operator switch { "equ" => value == condition.Value, "neq" => value != condition.Value, _ => true };
                }
                case "cnt":
                {
                    int count = condition.Iterator.Axes.Count == 0 ? 0 : model.Along(condition.Iterator, WalkFrom, lastStepWhole: false).Count;
                    return Compare(condition.Operator, count, Int(condition.Value, 0));
                }
                case "maxDepth": return Compare(condition.Operator, model.MaxDepth(WalkFrom), Int(condition.Value, 0));
                case "depth": return Compare(condition.Operator, model.Depth(WalkFrom), Int(condition.Value, 0));
                case "pos" or "revPos" or "posEven" or "posOdd":
                {
                    if (model.ParentOf(WalkFrom, out _) is not { } parent) return true;
                    List<string> siblings = model.ChildrenOf(parent, "node");
                    int position = siblings.IndexOf(WalkFrom) + 1;
                    if (position <= 0) return true;
                    return condition.Function switch
                    {
                        "pos" => Compare(condition.Operator, position, Int(condition.Value, 0)),
                        "revPos" => Compare(condition.Operator, siblings.Count - position + 1, Int(condition.Value, 0)),
                        "posEven" => Compare(condition.Operator, position % 2 == 0 ? 1 : 0, Int(condition.Value, 1)),
                        _ => Compare(condition.Operator, position % 2 == 1 ? 1 : 0, Int(condition.Value, 1)),
                    };
                }
            }
            return true;
        }

        private string? HierarchyBranchAbove(Point point)
        {
            Connection? up = Diagram.Model.Connections.FirstOrDefault(c => c.Type == "presParOf" && c.Destination == point.Id);
            return up is not null && Diagram.Model.ById.TryGetValue(up.Source, out Point? parent) ? parent.HierarchyBranch : null;
        }

        private static bool Compare(string op, int first, int second) => op switch
        {
            "equ" => first == second, "gt" => first > second, "gte" => first >= second,
            "lt" => first < second, "lte" => first <= second, "neq" => first != second, _ => false,
        };

        /// <summary>visit(ForEachAtom): one pass per data point the axis reaches from the point the enclosing pass stands on.</summary>
        private void VisitForEach(ForEachAtom forEach)
        {
            if (forEach.Reference.Length > 0)
            {
                // a ref already being followed (itself, an ancestor, a loop of refs) is not followed again
                if (Diagram.Layout.ForEachByName.TryGetValue(forEach.Reference, out ForEachAtom? referenced) && _followedReferences.Add(referenced))
                {
                    VisitForEach(referenced);
                    _followedReferences.Remove(referenced);
                }
                return;
            }
            Iterator iterator = forEach.Iterator;
            DataModel model = Diagram.Model;
            bool overTransitions = iterator.PointType is "sibTrans" or "parTrans" or "all";
            if (iterator.HideLastTransition && iterator.Axes.Count > 0 && iterator.Axes[0] == "followSib" && overTransitions)
            {
                // the transition after the last node is hidden: no node follows the one the pass stands on
                var nextNode = new Iterator { Axes = ["followSib"], PointTypes = ["node"] };
                if (WalkFrom.Length > 0 && model.Along(nextNode, WalkFrom, lastStepWhole: true).Count == 0) return;
            }
            bool alongSiblings = iterator.Axes.Count == 0 || iterator.Axes[0] is "followSib" or "precedSib" or "self";
            if (iterator.PointType == "sibTrans" && alongSiblings)
            {
                // the connector after the node the pass stands on: the sibTrans of that node
                string held = PassNodeId;
                string transition = "";
                if (held.Length > 0)
                    transition = model.Connections.FirstOrDefault(c => c.Type == "parOf" && c.Destination == held && c.SiblingTransition.Length > 0)?.SiblingTransition ?? "";
                PassNodeId = transition;
                DefaultVisit(forEach);
                PassNodeId = held;
                return;
            }
            if (iterator.Axes.Count == 1 && iterator.Axes[0] == "self")
            {
                DefaultVisit(forEach);
                return;
            }
            string from = WalkFrom;
            List<string> passes = from.Length > 0 ? model.Along(iterator, from, lastStepWhole: true) : [];
            int step = iterator.Step == 0 ? 1 : iterator.Step;
            int startWanted = iterator.Start, countWanted = iterator.Count;
            if (iterator.Axes.Count > 1)
            {
                if (iterator.Starts.Count >= iterator.Axes.Count) startWanted = iterator.Starts[iterator.Starts.Count - 1];
                if (iterator.Counts.Count >= iterator.Axes.Count) countWanted = iterator.Counts[iterator.Counts.Count - 1] == 0 ? -1 : iterator.Counts[iterator.Counts.Count - 1];
            }
            int along = passes.Count;
            int firstAlong = step < 0 ? along - 1 : 0;
            if (startWanted > 0) firstAlong = startWanted - 1;
            else if (startWanted < 0) firstAlong = along + startWanted;
            int room = step < 0 ? firstAlong + 1 : along - firstAlong;
            int children = along > 0 ? Math.Max(room, 0) : from.Length > 0 ? 0 : 1;
            int count = countWanted == -1 ? children : Math.Min(children, countWanted);
            (string heldPass, int heldIndex, int heldStep, int heldCount) = (PassNodeId, CurrentIndex, CurrentStep, CurrentCount);
            CurrentStep = step;
            CurrentCount = count;
            int alongIndex = firstAlong;
            for (CurrentIndex = step < 0 ? count - 1 : 0; CurrentIndex >= 0 && CurrentIndex < count; CurrentIndex += step, alongIndex += step)
            {
                PassNodeId = alongIndex >= 0 && alongIndex < along ? passes[alongIndex] : "";
                DefaultVisit(forEach);
            }
            (PassNodeId, CurrentIndex, CurrentStep, CurrentCount) = (heldPass, heldIndex, heldStep, heldCount);
        }

        /// <summary>presentationForPass: the presentation point of a layout node's name under the current one, built for the pass's data point.</summary>
        protected Point? PresentationFor(LayoutNodeAtom node)
        {
            DataModel model = Diagram.Model;
            List<Point> under = model.Points.Where(p => p.PresName == node.Name && p.Type == "pres" && model.IsConnected(Current.Id, p.Id)).ToList();
            if (PassNodeId.Length > 0)
            {
                if (under.FirstOrDefault(p => p.PresAssocId == PassNodeId) is { } own) return own;
                Connection? mine = model.Connections.FirstOrDefault(c => c.Type == "parOf" && c.Destination == PassNodeId);
                if (mine is not null && mine.SourceOrder > 0)
                {
                    Connection? before = model.Connections.FirstOrDefault(c => c.Type == "parOf" && c.Source == mine.Source && c.SourceOrder == mine.SourceOrder - 1 && c.SiblingTransition.Length > 0);
                    if (before is not null && under.FirstOrDefault(p => p.PresAssocId == before.SiblingTransition) is { } transition) return transition;
                }
                return null;
            }
            if (CurrentIndex >= 0 && CurrentIndex < under.Count && CurrentCount > 0) return under[CurrentIndex];
            return under.Count >= 1 ? under[0] : null;
        }

        protected void Descend(Point presentation, Action body)
        {
            (Point heldCurrent, string heldPass) = (Current, PassNodeId);
            Current = presentation;
            PassNodeId = "";
            body();
            (Current, PassNodeId) = (heldCurrent, heldPass);
        }
    }

    /// <summary>Counts nothing; finds the shape template of a layout node (ShapeTemplateVisitor: only choose / if are entered).</summary>
    private sealed class ShapeTemplateVisitor(Diagram diagram, Point current) : Visitor(diagram, current)
    {
        public ShapeAtom? Template;
        public readonly List<Constraint> Constraints = [];
        public void VisitBody(LayoutNodeAtom node) => DefaultVisit(node);
        protected override void VisitShape(ShapeAtom shape) => Template ??= shape;
        protected override void VisitConstraint(ConstraintAtom constraint) => Constraints.Add(constraint.Constraint);
    }

    /// <summary>ShapeCreationVisitor: one shape per layout node a presentation point is built for, nested as the layout nests.</summary>
    private sealed class ShapeCreationVisitor : Visitor
    {
        private DiagramShape _parent;
        private readonly LayoutNodeAtom _root;
        private bool _lookingForAlgorithm;

        public ShapeCreationVisitor(Diagram diagram, Point root, DiagramShape frame, LayoutNodeAtom rootNode) : base(diagram, root)
            => (_parent, _root) = (frame, rootNode);

        protected override void VisitAlgorithm(AlgorithmAtom algorithm)
        {
            if (!_lookingForAlgorithm) return;
            _parent.AspectRatio = algorithm.AspectRatio;
            _parent.AlgorithmType = algorithm.Type;
            _parent.VerticalShapesCount = VerticalShapesCount(algorithm, _parent);
        }

        protected override void VisitLayoutNode(LayoutNodeAtom node)
        {
            if (_lookingForAlgorithm) return;
            if (PresentationFor(node) is not { } presentation) return;
            DiagramShape parent = _parent, shapeParent = _parent;
            if (node == _root)
            {
                SetUpShape(_parent, node, presentation);
                Diagram.ShapeOfPresentation[presentation] = _parent;
            }
            else
            {
                var template = new ShapeTemplateVisitor(Diagram, presentation);
                template.VisitBody(node);
                if (template.Template is { } atom)
                {
                    var shape = new DiagramShape
                    {
                        Name = node.Name, IsGroup = atom.Type is null, Preset = atom.HideGeometry ? null : atom.Type, HideGeometry = atom.HideGeometry || atom.Type is "none" or "",
                        Rotation = atom.Rotation, ZOrderOffset = atom.ZOrderOffset, Adjustments = atom.Adjustments,
                    };
                    shape.OwnConstraints = template.Constraints;
                    SetUpShape(shape, node, presentation);
                    _parent.Children.Add(shape);
                    shapeParent = shape;
                    Diagram.ShapeOfPresentation[presentation] = shape;
                }
            }
            Descend(presentation, () =>
            {
                _parent = shapeParent;
                DefaultVisit(node);
                _lookingForAlgorithm = true;
                DefaultVisit(node);
                _lookingForAlgorithm = false;
                _parent = parent;
            });
        }

        /// <summary>LayoutNode::setupShape: the data points' text (presOf, with their outline level), their shape properties, the style label.</summary>
        private void SetUpShape(DiagramShape shape, LayoutNodeAtom node, Point presentation)
        {
            DataModel model = Diagram.Model;
            shape.Presentation = presentation;
            var sources = model.Connections.Where(c => c.Type == "presOf" && c.Destination == presentation.Id).OrderBy(c => c.DestinationOrder).ToList();
            if (sources.Count > 0)
            {
                var depths = sources.Select(c => model.Depth(c.Source)).ToList();
                int minimum = depths.Min();
                for (int i = 0; i < sources.Count; i++)
                {
                    if (!model.ById.TryGetValue(sources[i].Source, out Point? data)) continue;
                    shape.DataNodeType = data.Type;
                    if (depths[i] == minimum && data.ShapeProperties is { } own) shape.OwnProperties.AddRange(own.Elements());
                    if (data.Text is null) continue;
                    foreach (XElement paragraph in data.Text.Elements("p"))
                        if (paragraph.Elements("r").Any() || paragraph.Elements("fld").Any()) shape.Paragraphs.Add((paragraph, depths[i]));
                }
            }
            else
                shape.DataNodeType = model.ById.TryGetValue(presentation.PresAssocId, out Point? associated) ? associated.Type : "";
            if (presentation.ShapeProperties is { } presentationProperties) shape.OwnProperties.AddRange(presentationProperties.Elements());
            shape.StyleLabel = presentation.PresStyleLabel.Length > 0 ? presentation.PresStyleLabel : node.StyleLabel;
            shape.ColorIndex = presentation.PresStyleIndex >= 0 ? presentation.PresStyleIndex : CurrentIndex;
        }
    }

    /// <summary>ShapeLayoutingVisitor: per layout node its constraints, rules and algorithm, then the nested nodes; parent constraints stay in force below.</summary>
    private sealed class ShapeLayoutingVisitor(Diagram diagram, Point root) : Visitor(diagram, root)
    {
        private enum Looking { Constraint, Rule, Algorithm, LayoutNode }
        private Looking _looking = Looking.LayoutNode;
        private readonly List<Constraint> _constraints = [];
        private readonly List<RuleAtom> _rules = [];
        private int _ownFrom;
        private readonly List<Constraint> _allConstraints = [];      // every constraint, unfiltered: the sizes of lin / snake children read them all

        protected override void VisitConstraint(ConstraintAtom constraint)
        {
            if (_looking != Looking.Constraint) return;
            ParseConstraint(constraint.Constraint, _constraints, requireForName: true);
            _allConstraints.Add(constraint.Constraint);
        }

        protected override void VisitRule(RuleAtom rule)
        {
            if (_looking == Looking.Rule && rule.ForName.Length > 0) _rules.Add(rule);
        }

        protected override void VisitAlgorithm(AlgorithmAtom algorithm)
        {
            if (_looking == Looking.Algorithm && Diagram.ShapeOfPresentation.TryGetValue(Current, out DiagramShape? shape))
                Algorithms.Lay(algorithm, shape, _constraints, _allConstraints.Skip(_ownFrom).ToList(), _rules, Diagram);
        }

        protected override void VisitLayoutNode(LayoutNodeAtom node)
        {
            if (_looking != Looking.LayoutNode) return;
            if (PresentationFor(node) is not { } presentation) return;
            int inherited = _constraints.Count, inheritedAll = _allConstraints.Count, heldOwnFrom = _ownFrom;
            _ownFrom = inheritedAll;
            Descend(presentation, () =>
            {
                foreach (Looking pass in new[] { Looking.Constraint, Looking.Rule, Looking.Algorithm, Looking.LayoutNode })
                {
                    _looking = pass;
                    DefaultVisit(node);
                }
            });
            _constraints.RemoveRange(inherited, _constraints.Count - inherited);
            _allConstraints.RemoveRange(inheritedAll, _allConstraints.Count - inheritedAll);
            _ownFrom = heldOwnFrom;
        }
    }

    /// <summary>ConstraintAtom::parseConstraint: equality constraints only; a for-name is needed except for spacing, margins and sibling transitions.</summary>
    private static void ParseConstraint(Constraint constraint, List<Constraint> into, bool requireForName)
    {
        if (requireForName && (constraint.Type is "sp" or "lMarg" or "rMarg" or "tMarg" or "bMarg" || constraint.PointType == "sibTrans")) requireForName = false;
        if (requireForName && constraint.ForName.Length == 0) return;
        if (constraint.Operator is "none" or "equ" && constraint.Type != "none") into.Add(constraint);
    }

    /// <summary>AlgAtom::getVerticalShapesCount.</summary>
    private static int VerticalShapesCount(AlgorithmAtom algorithm, DiagramShape shape)
    {
        if (shape.Children.Count == 0) return shape.Preset != "conn" ? 1 : 0;
        string direction = algorithm.Type == "hierRoot" ? "fromT" : algorithm.Parameter("linDir", "fromL");
        string secondary = algorithm.Parameter("secLinDir", "");
        int count = 0;
        if (direction is "fromT" or "fromB")
            foreach (DiagramShape child in shape.Children) count += child.VerticalShapesCount;
        else if (direction is "fromL" or "fromR" && secondary == "fromT")
        {
            foreach (DiagramShape child in shape.Children) count += child.VerticalShapesCount;
            count = (count + 1) / 2;
        }
        else
            foreach (DiagramShape child in shape.Children) count = Math.Max(count, child.VerticalShapesCount);
        return count;
    }

    // ── the algorithms (AlgAtom::layoutShape) ─────────────────────────────────────────────────────────────────────────────

    private static class Algorithms
    {
        public static void Lay(AlgorithmAtom algorithm, DiagramShape shape, List<Constraint> constraints, List<Constraint> own, List<RuleAtom> rules, Diagram diagram)
        {
            if (algorithm.Type != "lin") shape.Children.RemoveAll(c => c.IsGroup && c.Children.Count == 0);
            switch (algorithm.Type)
            {
                case "composite": Composite(algorithm, shape, constraints); break;
                case "conn": Connector(algorithm, shape); break;
                case "cycle": Cycle(algorithm, shape, own); break;
                case "hierChild" or "hierRoot": Hierarchy(algorithm, shape, own); break;
                case "lin": Linear(algorithm, shape, own); break;
                case "pyra": Pyramid(algorithm, shape); break;
                case "snake": Snake(algorithm, shape, own); break;
                case "sp": shape.Paragraphs.Clear(); break;
                case "tx": Text(algorithm, shape, constraints, rules); break;
            }
        }

        private static bool IsFontUnit(string type) => type is "primFontSz" or "secFontSz";

        private static bool InferFromProperty(Dictionary<string, double> map, string referenceType, out double value)
        {
            value = 0;
            if (referenceType == "r" && map.TryGetValue("l", out double left) && map.TryGetValue("w", out double width))
            {
                value = left + width;
                return true;
            }
            return false;
        }

        /// <summary>CompositeAlg::applyConstraintToLayout.</summary>
        private static void ApplyConstraint(Constraint constraint, Dictionary<string, Dictionary<string, double>> properties)
        {
            if (constraint.ForName.Length == 0) return;
            if (!properties.TryGetValue(constraint.RefForName, out Dictionary<string, double>? reference)) return;
            if (!properties.TryGetValue(constraint.ForName, out Dictionary<string, double>? target)) properties[constraint.ForName] = target = [];
            if (reference.TryGetValue(constraint.RefType, out double value)) target[constraint.Type] = value * constraint.Factor;
            else if (InferFromProperty(reference, constraint.RefType, out double inferred)) target[constraint.Type] = inferred * constraint.Factor;
            else target[constraint.Type] = constraint.Value * (IsFontUnit(constraint.RefType) ? EmuPerPoint : EmuPerMillimetre);
        }

        /// <summary>CompositeAlg::layoutShapeChildren: children placed by their constraints (l / t / r / b / ctrX / ctrY / w / h), content centred vertically.</summary>
        private static void Composite(AlgorithmAtom algorithm, DiagramShape shape, List<Constraint> constraints)
        {
            var properties = new Dictionary<string, Dictionary<string, double>>();
            var parent = properties[""] = [];
            double parentXOffset = 0;
            if (algorithm.AspectRatio != 1.0)
            {
                parent["w"] = shape.Width; parent["h"] = shape.Height; parent["l"] = 0; parent["t"] = 0; parent["r"] = shape.Width; parent["b"] = shape.Height;
            }
            else
            {
                parent["w"] = Math.Min(shape.Width, shape.Height);
                parent["h"] = shape.Height;
                if (parent["w"] < shape.Width) parentXOffset = (shape.Width - parent["w"]) / 2;
                parent["l"] = parentXOffset; parent["t"] = 0; parent["r"] = shape.Width - parentXOffset; parent["b"] = shape.Height;
            }
            foreach (Constraint constraint in constraints) ApplyConstraint(constraint, properties);
            double verticalMin = double.MaxValue, verticalMax = 0;
            LayoutNodeAtom? ownNode = algorithm.OwnerNode;
            foreach (DiagramShape child in shape.Children)
            {
                foreach (Constraint constraint in constraints.Where(c => c.ForName == child.Name)) ApplyConstraint(constraint, properties);
                foreach (LayoutNodeAtom childNode in ownNode?.Children.OfType<LayoutNodeAtom>().Where(n => n.Name == child.Name) ?? [])
                    foreach (ConstraintAtom own in childNode.Children.OfType<ConstraintAtom>())
                    {
                        Constraint constraint = own.Constraint;
                        if (constraint.ForName.Length > 0 || constraint.RefForName.Length > 0) continue;
                        if (constraint.Value == 0 && constraint.RefType == "none") continue;
                        Constraint copy = constraint.Copy();
                        copy.ForName = childNode.Name;
                        copy.RefForName = childNode.Name;
                        ApplyConstraint(copy, properties);
                    }
                double width = shape.Width, height = shape.Height, x = 0, y = 0;
                if (child.FillsParent) { (child.X, child.Y, child.Width, child.Height) = (0, 0, shape.Width, shape.Height); continue; }
                if (properties.TryGetValue(child.Name, out Dictionary<string, double>? own2))
                {
                    if (own2.TryGetValue("w", out double w)) width = Math.Min(w, shape.Width);
                    if (own2.TryGetValue("h", out double h)) height = Math.Min(h, shape.Height);
                    if (own2.TryGetValue("l", out double l)) x = l;
                    else if (own2.TryGetValue("ctrX", out double cx)) x = cx - width / 2;
                    else if (own2.TryGetValue("r", out double r)) x = r - width;
                    if (own2.TryGetValue("t", out double t)) y = t;
                    else if (own2.TryGetValue("ctrY", out double cy)) y = cy - height / 2;
                    else if (own2.TryGetValue("b", out double b)) y = b - height;
                    if (own2.TryGetValue("l", out double l2) && own2.TryGetValue("r", out double r2)) width = r2 - l2;
                    if (own2.TryGetValue("t", out double t2) && own2.TryGetValue("b", out double b2)) height = b2 - t2;
                    x += parentXOffset;
                    width = Math.Min(width, shape.Width - x);
                    height = Math.Min(height, shape.Height - y);
                }
                (child.X, child.Y, child.Width, child.Height) = (x, y, width, height);
                verticalMin = Math.Min(y, verticalMin);
                verticalMax = Math.Max(y + height, verticalMax);
            }
            if (!(verticalMin >= 0 && verticalMin <= verticalMax && verticalMax <= parent["h"])) return;
            double difference = parent["h"] - (verticalMax - verticalMin);
            if (difference > 0)
                foreach (DiagramShape child in shape.Children) child.Y += difference / 2;
        }

        /// <summary>AlgAtom::getConnectorType: the arrow a conn shape becomes.</summary>
        private static string? ConnectorType(AlgorithmAtom algorithm)
        {
            string routing = algorithm.Parameter("connRout", ""), begin = algorithm.Parameter("begSty", ""), end = algorithm.Parameter("endSty", "");
            if (routing == "bend") return null;     // org chart connectors: LibreOffice hides them (getConnectorType returns 0)
            if (begin == "arr" && end == "arr") return "leftRightArrow";
            if (begin == "arr") return "leftArrow";
            return "rightArrow";
        }

        private static void Connector(AlgorithmAtom algorithm, DiagramShape shape)
        {
            if (shape.Preset == "conn")
            {
                shape.Preset = ConnectorType(algorithm);
                if (shape.Preset is null) shape.HideGeometry = true;
            }
            // LibreOffice resizes by the direct constraints but then keeps the parent's own size (conn: rParent w / h), so the shape keeps its place
        }

        /// <summary>
        /// The cycle algorithm (ECMA-376 Part 1 21.4.2 "Cycle"; Word's drawing for the built-in cycles): the nodes on a circle
        /// from stAng (0 = top, clockwise) over spanAng, their centres one node width plus sibSp apart along the chord, so the
        /// radius is (w + sibSp) / (2 sin(step / 2)); the ring of nodes scaled to fit the shape and centred. A connector
        /// (sibTrans) sits on the chord's middle between its node and the next, turned along it, sized by its constraints.
        /// </summary>
        private static void Cycle(AlgorithmAtom algorithm, DiagramShape shape, List<Constraint> own)
        {
            if (shape.Children.Count == 0 || shape.Width == 0 || shape.Height == 0) return;
            double startAngle = algorithm.Number("stAng", 0), spanAngle = algorithm.Number("spanAng", 360);
            List<Dictionary<string, double>> sizes = ChildSizes(shape, own, out _);
            Dictionary<string, Dictionary<string, double>> keyed = Resolve(own, shape.Width, shape.Height);
            var nodes = new List<int>();
            for (int i = 0; i < shape.Children.Count; i++)
                if (!IsConnector(shape.Children[i])) nodes.Add(i);
            int count = nodes.Count;
            if (count == 0) return;
            DiagramShape? centre = null;
            if (algorithm.Parameter("ctrShpMap", "none") == "fNode" && count > 1)
            {
                // the first node in the centre, the others around it
                centre = shape.Children[nodes[0]];
                nodes.RemoveAt(0);
                count--;
            }
            Dictionary<string, double> nodeSize = sizes[nodes[0]];
            double width = nodeSize.TryGetValue("w", out double w) ? w : shape.Width / 4, height = nodeSize.TryGetValue("h", out double h) ? h : width;
            double spacing = keyed["self"].TryGetValue("sibSp", out double sibSp) ? sibSp : width * 0.5;
            double step = spanAngle / count * Math.PI / 180;
            double half = Math.Abs(Math.Sin(step / 2));
            double radius = count == 1 || half < 1e-9 ? 0 : (width + spacing) / (2 * half);
            var centres = new (double X, double Y)[count];
            double left = double.MaxValue, top = double.MaxValue, right = double.MinValue, bottom = double.MinValue;
            for (int k = 0; k < count; k++)
            {
                double angle = startAngle * Math.PI / 180 + k * step;
                centres[k] = (radius * Math.Sin(angle), -radius * Math.Cos(angle));
                left = Math.Min(left, centres[k].X - width / 2); right = Math.Max(right, centres[k].X + width / 2);
                top = Math.Min(top, centres[k].Y - height / 2); bottom = Math.Max(bottom, centres[k].Y + height / 2);
            }
            double scale = right - left > 0 && bottom - top > 0 ? Math.Min(shape.Width / (right - left), shape.Height / (bottom - top)) : 1;
            double offsetX = (shape.Width - (right - left) * scale) / 2 - left * scale, offsetY = (shape.Height - (bottom - top) * scale) / 2 - top * scale;
            for (int k = 0; k < count; k++)
            {
                DiagramShape node = shape.Children[nodes[k]];
                (node.Width, node.Height) = (width * scale, height * scale);
                (node.X, node.Y) = (offsetX + centres[k].X * scale - node.Width / 2, offsetY + centres[k].Y * scale - node.Height / 2);
            }
            if (centre is not null)
            {
                (centre.Width, centre.Height) = (width * scale, height * scale);
                (centre.X, centre.Y) = (offsetX - centre.Width / 2, offsetY - centre.Height / 2);
            }
            // each connector follows its node (hideLastTrans="0" keeps the one from the last node back to the first)
            int nodeIndex = -1;
            for (int i = 0; i < shape.Children.Count; i++)
            {
                DiagramShape child = shape.Children[i];
                int found = nodes.IndexOf(i);
                if (found >= 0) { nodeIndex = found; continue; }
                if (!IsConnector(child) || nodeIndex < 0 || count < 2) continue;
                var from = centres[nodeIndex];
                var to = centres[(nodeIndex + 1) % count];
                double connectorWidth = (sizes[i].TryGetValue("w", out double cw) ? cw : width * 0.25) * scale;
                double connectorHeight = connectorWidth;
                if (child.OwnConstraints.FirstOrDefault(c => c.Type == "h" && c.RefType == "w" && c.ForName.Length == 0 && c.For == "none") is { } derived) connectorHeight = connectorWidth * derived.Factor;
                double middleX = offsetX + (from.X + to.X) / 2 * scale, middleY = offsetY + (from.Y + to.Y) / 2 * scale;
                (child.Width, child.Height) = (connectorWidth, connectorHeight);
                (child.X, child.Y) = (middleX - connectorWidth / 2, middleY - connectorHeight / 2);
                child.Rotation = Math.Atan2(to.Y - from.Y, to.X - from.X) * 180 / Math.PI;
                if (child.Rotation < 0) child.Rotation += 360;
            }
        }

        private static bool IsConnector(DiagramShape shape) => shape.Preset == "conn" || shape.DataNodeType is "sibTrans" or "parTrans";

        /// <summary>
        /// hierRoot / hierChild (ECMA-376 Part 1 21.4.2 "Hierarchy Root" / "Hierarchy Child"; Word's drawings of the built-in org
        /// chart and hierarchies): the whole tree is laid out from the outermost hierarchy node at once. A hierRoot is its node's
        /// box (the composite the constraints size) above the rows of its hierChild groups, centred over them, sp below it; a
        /// hierChild is a row of hierRoots sibSp apart (linDir fromL / fromR). The tree is scaled to fit the shape and centred;
        /// the connectors are not drawn (Word's elbow lines; LibreOffice hides them as well).
        /// </summary>
        private static void Hierarchy(AlgorithmAtom algorithm, DiagramShape shape, List<Constraint> own)
        {
            if (shape.HierarchyPlaced || shape.Children.Count == 0 || shape.Width == 0 || shape.Height == 0) return;
            Dictionary<string, Dictionary<string, double>> keyed = Resolve(own, shape.Width, shape.Height);
            (double Width, double Height)? anyBox = null;
            foreach (var pair in keyed)
                if (pair.Key.StartsWith("n:", StringComparison.Ordinal) && pair.Value.TryGetValue("w", out double bw) && pair.Value.TryGetValue("h", out double bh)) { anyBox = (bw, bh); break; }
            (double boxWidth, double boxHeight) = anyBox ?? (shape.Width / 4, shape.Width / 8);
            double siblingSpace = keyed["self"].TryGetValue("sibSp", out double sib) ? sib : boxWidth * 0.1;
            double defaultGap = keyed["self"].TryGetValue("sp", out double sp0) ? sp0 : boxHeight * 0.25;
            var measured = new Dictionary<DiagramShape, (double Width, double Height)>();

            (double, double) Box(DiagramShape box)
                => keyed.TryGetValue("n:" + box.Name, out var own2) && own2.TryGetValue("w", out double w) && own2.TryGetValue("h", out double h) ? (w, h) : (boxWidth, boxHeight);
            double Gap(DiagramShape root) => keyed.TryGetValue("n:" + root.Name, out var own2) && own2.TryGetValue("sp", out double gap) ? gap : defaultGap;
            bool IsRow(DiagramShape s) => s.AlgorithmType == "hierChild";
            bool IsRoot(DiagramShape s) => s.AlgorithmType == "hierRoot";

            (double Width, double Height) Measure(DiagramShape node, int depth)
            {
                if (depth > 64) return (0, 0);
                (double Width, double Height) size;
                if (IsRow(node))
                {
                    var items = node.Children.Where(IsRoot).Select(c => Measure(c, depth + 1)).ToList();
                    size = items.Count == 0 ? (0, 0) : (items.Sum(i => i.Width) + siblingSpace * (items.Count - 1), items.Max(i => i.Height));
                }
                else
                {
                    DiagramShape? box = node.Children.FirstOrDefault(c => !IsRow(c) && !IsRoot(c) && !IsConnector(c));
                    (double width, double height) = box is null ? (0.0, 0.0) : Box(box);
                    double gap = Gap(node);
                    foreach (DiagramShape row in node.Children.Where(IsRow))
                    {
                        var rowSize = Measure(row, depth + 1);
                        if (rowSize.Width <= 0) continue;
                        width = Math.Max(width, rowSize.Width);
                        height += gap + rowSize.Height;
                    }
                    size = (width, height);
                }
                measured[node] = size;
                return size;
            }

            void Place(DiagramShape node, double scale, int depth)
            {
                if (depth > 64) return;
                node.HierarchyPlaced = true;
                (double width, double _) = measured[node];
                if (IsRow(node))
                {
                    bool fromRight = algorithm.Parameter("linDir", "fromL") == "fromR" && node == shape;
                    double x = 0;
                    foreach (DiagramShape item in node.Children.Where(IsRoot))
                    {
                        var itemSize = measured[item];
                        (item.X, item.Y, item.Width, item.Height) = (fromRight ? width * scale - x - itemSize.Width * scale : x, 0, itemSize.Width * scale, itemSize.Height * scale);
                        Place(item, scale, depth + 1);
                        x += (itemSize.Width + siblingSpace) * scale;
                    }
                }
                else
                {
                    double y = 0;
                    if (node.Children.FirstOrDefault(c => !IsRow(c) && !IsRoot(c) && !IsConnector(c)) is { } box)
                    {
                        (double bw, double bh) = Box(box);
                        (box.X, box.Y, box.Width, box.Height) = ((width - bw) / 2 * scale, 0, bw * scale, bh * scale);
                        y = bh;
                    }
                    double gap = Gap(node);
                    foreach (DiagramShape row in node.Children.Where(IsRow))
                    {
                        var rowSize = measured[row];
                        if (rowSize.Width <= 0) { row.HierarchyPlaced = true; continue; }
                        y += gap;
                        (row.X, row.Y, row.Width, row.Height) = ((width - rowSize.Width) / 2 * scale, y * scale, rowSize.Width * scale, rowSize.Height * scale);
                        Place(row, scale, depth + 1);
                        y += rowSize.Height;
                    }
                }
                foreach (DiagramShape connector in node.Children.Where(IsConnector)) { connector.HideGeometry = true; connector.Width = connector.Height = 0; }
            }

            (double totalWidth, double totalHeight) = Measure(shape, 0);
            if (totalWidth <= 0 || totalHeight <= 0) return;
            double fit = Math.Min(shape.Width / totalWidth, shape.Height / totalHeight);
            Place(shape, fit, 0);
            // the tree centred in the shape
            double dx = (shape.Width - totalWidth * fit) / 2, dy = (shape.Height - totalHeight * fit) / 2;
            foreach (DiagramShape child in shape.Children) { child.X += dx; child.Y += dy; }
        }

        /// <summary>
        /// The children's sizes from the layout node's own constraints, in their order (ECMA-376 Part 1 21.4.4 / 21.4.7.10): a
        /// constraint sets the w / h / sp of the children it names (forName, else its ptType, else all of them) to a reference
        /// times its factor: the parent's own size (no refFor), or what an earlier constraint gave other children (refForName /
        /// refPtType), a font size giving millimetres (as Word reads primFontSz-based heights); without a reference its value (mm).
        /// Equalities without a value or reference only tie children together and set nothing. Then each child's own self
        /// constraints (h = 0.6 w and the like). Unset sizes stay null.
        /// </summary>
        /// <summary>
        /// A layout node's own constraints resolved in their order, again until settled (a reference stated later, as a height
        /// from the font size given after it): values keyed "n:" + forName, "t:" + ptType, "all" (every child) or "self" (the
        /// node itself: its w / h to start with, its sp / sibSp / secSibSp); a reference without refFor reads "self".
        /// </summary>
        private static Dictionary<string, Dictionary<string, double>> Resolve(List<Constraint> own, double width, double height)
        {
            var keyed = new Dictionary<string, Dictionary<string, double>> { ["self"] = new() { ["w"] = width, ["h"] = height } };
            for (int pass = 0; pass < 3; pass++)
            foreach (Constraint constraint in own)
            {
                if (constraint.Operator is not ("none" or "equ")) continue;
                double? value = null;
                if (constraint.RefType == "none") { if (constraint.Value != 0) value = constraint.Value * (IsFontUnit(constraint.Type) ? 1 : EmuPerMillimetre); }
                else
                {
                    Dictionary<string, double>? reference = constraint.RefForName.Length > 0 ? Get(keyed, "n:" + constraint.RefForName)
                                                          : constraint.RefPointType() is { } refType ? Get(keyed, "t:" + refType)
                                                          : constraint.RefFor is "none" or "self" ? keyed["self"] : null;
                    if (reference is not null && reference.TryGetValue(constraint.RefType, out double referenced))
                        value = referenced * constraint.Factor * (IsFontUnit(constraint.RefType) && !IsFontUnit(constraint.Type) ? EmuPerMillimetre : 1);
                }
                if (value is null) continue;
                string key = constraint.ForName.Length > 0 ? "n:" + constraint.ForName : constraint.PointType != "none" ? "t:" + constraint.PointType
                           : constraint.For is "ch" or "des" ? "all" : "self";
                if (key == "self" && constraint.Type is "w" or "h") continue;     // the node's own size is its parent's to set
                if (!keyed.TryGetValue(key, out Dictionary<string, double>? target)) keyed[key] = target = [];
                target[constraint.Type] = value.Value;
            }
            return keyed;
        }

        private static List<Dictionary<string, double>> ChildSizes(DiagramShape shape, List<Constraint> own, out double? spacing)
        {
            Dictionary<string, Dictionary<string, double>> keyed = Resolve(own, shape.Width, shape.Height);
            spacing = keyed["self"].TryGetValue("sp", out double space) ? space : null;
            var sizes = new List<Dictionary<string, double>>();
            foreach (DiagramShape child in shape.Children)
            {
                var size = new Dictionary<string, double>();
                if (keyed.TryGetValue("all", out var all)) foreach (var pair in all) size[pair.Key] = pair.Value;
                foreach (var pair in keyed.Where(k => k.Key.StartsWith("t:", StringComparison.Ordinal) && DataModel.IsOfKind(child.DataNodeType, k.Key.Substring(2))))
                    foreach (var value in pair.Value) size[value.Key] = value.Value;
                if (keyed.TryGetValue("n:" + child.Name, out var named)) foreach (var pair in named) size[pair.Key] = pair.Value;
                foreach (Constraint constraint in child.OwnConstraints)
                {
                    if (constraint.ForName.Length > 0 || constraint.RefForName.Length > 0 || constraint.Type is not ("w" or "h") || constraint.Operator is not ("none" or "equ")) continue;
                    if (constraint.RefType == "none") { if (constraint.Value != 0) size[constraint.Type] = constraint.Value * EmuPerMillimetre; }
                    else if (size.TryGetValue(constraint.RefType, out double referenced)) size[constraint.Type] = referenced * constraint.Factor;
                }
                if (child.OwnConstraints.FirstOrDefault(c => c.Type == "begPad") is { } begin) child.BeginPad = begin.Factor;
                if (child.OwnConstraints.FirstOrDefault(c => c.Type == "endPad") is { } end) child.EndPad = end.Factor;
                sizes.Add(size);
            }
            return sizes;
        }

        private static Dictionary<string, double>? Get(Dictionary<string, Dictionary<string, double>> keyed, string key) => keyed.TryGetValue(key, out var found) ? found : null;

        /// <summary>
        /// The lin algorithm (ECMA-376 Part 1 21.4.2 "Linear Flow"): the children one after another along linDir at the sizes
        /// their constraints give, then all of them scaled by one factor so the row fits the shape both ways, centred in it (the
        /// across-axis alignment of vertAlign / horzAlign mid). A connector child keeps its space and is drawn within it, between
        /// its begin and end padding, turned to the flow. chOrder="t" puts the later children behind the earlier ones.
        /// </summary>
        private static void Linear(AlgorithmAtom algorithm, DiagramShape shape, List<Constraint> own)
        {
            if (shape.Children.Count == 0 || shape.Width == 0 || shape.Height == 0) return;
            string direction = algorithm.Parameter("linDir", "fromL");
            bool horizontal = direction is "fromL" or "fromR", reverse = direction is "fromR" or "fromB";
            List<Dictionary<string, double>> sizes = ChildSizes(shape, own, out double? spacing);
            int count = shape.Children.Count;
            double along = 0, across = 0;
            var lengths = new double[count];
            var breadths = new double[count];
            for (int i = 0; i < count; i++)
            {
                double fallbackAlong = (horizontal ? shape.Width : shape.Height) / count;
                lengths[i] = sizes[i].TryGetValue(horizontal ? "w" : "h", out double length) ? length : shape.Children[i].IsGroup && shape.Children[i].Children.Count == 0 ? 0 : fallbackAlong;
                breadths[i] = sizes[i].TryGetValue(horizontal ? "h" : "w", out double breadth) ? breadth : horizontal ? shape.Height : shape.Width;
                along += lengths[i];
                across = Math.Max(across, breadths[i]);
            }
            along += (spacing ?? 0) * Math.Max(count - 1, 0);
            double alongRoom = horizontal ? shape.Width : shape.Height, acrossRoom = horizontal ? shape.Height : shape.Width;
            // the row is scaled along the flow to fit; a size across that a child's own constraint derives from its length along
            // (h = 0.6 w) follows the scaled length, any other keeps its constrained size; both stay inside the shape
            string alongType = horizontal ? "w" : "h", acrossType = horizontal ? "h" : "w";
            double scale = Math.Min(1, along > 0 ? alongRoom / along : 1);
            double position = (alongRoom - along * scale) / 2;
            for (int k = 0; k < count; k++)
            {
                int i = reverse ? count - 1 - k : k;
                DiagramShape child = shape.Children[i];
                double length = lengths[i] * scale, breadth = breadths[i];
                if (child.OwnConstraints.FirstOrDefault(c => c.Type == acrossType && c.RefType == alongType && c.ForName.Length == 0 && c.Operator is "none" or "equ") is { } derived)
                    breadth = length * derived.Factor;
                breadth = Math.Min(breadth, acrossRoom);
                double start = position;
                position += length + (spacing ?? 0) * scale;
                if (child.BeginPad > 0 || child.EndPad > 0)
                {
                    // a connector is drawn within its space, between the padding at its two ends
                    start += length * child.BeginPad;
                    length *= 1 - child.BeginPad - child.EndPad;
                }
                double offset = (acrossRoom - breadth) / 2;
                if (horizontal) (child.X, child.Y, child.Width, child.Height) = (start, offset, length, breadth);
                else (child.X, child.Y, child.Width, child.Height) = (offset, start, breadth, length);
                if (child.Preset == "conn" || child.DataNodeType == "sibTrans" && child.Children.Count == 0)
                    child.Rotation = direction switch { "fromR" => 180, "fromT" => 90, "fromB" => 270, _ => 0 };
            }
            shape.Children.RemoveAll(c => c.IsGroup && c.Children.Count == 0);
            LayoutNodeAtom? parentNode = algorithm.Parent?.OwnerNode;
            if (parentNode?.ChildOrder == "t") shape.Children.Reverse();
        }

        /// <summary>
        /// The snake algorithm (ECMA-376 Part 1 21.4.2 "Snake"): the nodes in rows (flowDir row) from the grDir corner, as many
        /// columns as give the largest nodes at their constrained aspect ratio with the sp spacing between them; the grid centred
        /// in the shape and a short last row centred (off ctr).
        /// </summary>
        private static void Snake(AlgorithmAtom algorithm, DiagramShape shape, List<Constraint> own)
        {
            shape.Children.RemoveAll(c => c.IsGroup && c.Children.Count == 0);
            if (shape.Children.Count == 0 || shape.Width == 0 || shape.Height == 0) return;
            List<Dictionary<string, double>> sizes = ChildSizes(shape, own, out double? spacing);
            double nodeWidth = sizes[0].TryGetValue("w", out double w) ? w : shape.Width;
            double nodeHeight = sizes[0].TryGetValue("h", out double h) ? h : nodeWidth * 0.6;
            if (!(nodeWidth > 0 && nodeHeight > 0)) return;
            double gap = spacing ?? nodeWidth * 0.3;
            int count = shape.Children.Count;
            (int columns, double scale) best = (1, 0);
            for (int columns = 1; columns <= count; columns++)
            {
                int rows = (count + columns - 1) / columns;
                double scale = Math.Min(shape.Width / (columns * nodeWidth + (columns - 1) * gap), shape.Height / (rows * nodeHeight + (rows - 1) * gap));
                if (scale > best.scale + 1e-9) best = (columns, scale);
            }
            int bestColumns = best.columns, bestRows = (count + bestColumns - 1) / bestColumns;
            double s = best.scale, width = nodeWidth * s, height = nodeHeight * s, space = gap * s;
            double gridWidth = bestColumns * width + (bestColumns - 1) * space, gridHeight = bestRows * height + (bestRows - 1) * space;
            string gridDirection = algorithm.Parameter("grDir", "tL");
            bool fromRight = gridDirection is "tR" or "bR", fromBottom = gridDirection is "bL" or "bR";
            bool centreLastRow = algorithm.Parameter("off", "ctr") == "ctr";
            for (int i = 0; i < count; i++)
            {
                int row = i / bestColumns, column = i % bestColumns;
                int inRow = Math.Min(bestColumns, count - row * bestColumns);
                double rowWidth = inRow * width + (inRow - 1) * space;
                double left = (shape.Width - gridWidth) / 2 + (centreLastRow ? (gridWidth - rowWidth) / 2 : 0) + column * (width + space);
                if (fromRight) left = shape.Width - left - width;
                double top = (shape.Height - gridHeight) / 2 + row * (height + space);
                if (fromBottom) top = shape.Height - top - height;
                (shape.Children[i].X, shape.Children[i].Y, shape.Children[i].Width, shape.Children[i].Height) = (left, top, width, height);
            }
        }

        /// <summary>
        /// The pyra algorithm (ECMA-376 Part 1 21.4.2 "Pyramid"): one level per child, of equal height, the k-th of n from the top
        /// k/n of the shape's width and centred, so the levels' sides line up into a triangle, the first node on top as Word
        /// draws it; the level shape (pyraLvlNode) fills its group.
        /// </summary>
        private static void Pyramid(AlgorithmAtom algorithm, DiagramShape shape)
        {
            if (shape.Children.Count == 0 || shape.Width == 0 || shape.Height == 0) return;
            int count = shape.Children.Count;
            string levelNode = algorithm.Parameter("pyraLvlNode", "level");
            double height = shape.Height / count;
            for (int i = 0; i < count; i++)
            {
                DiagramShape child = shape.Children[i];
                double width = shape.Width * (i + 1) / count;
                (child.X, child.Y, child.Width, child.Height) = ((shape.Width - width) / 2, i * height, width, height);
                foreach (DiagramShape level in child.Children.Where(c => c.Name == levelNode)) level.FillsParent = true;
                if (child.Name == levelNode) child.FillsParent = false;
            }
        }

        /// <summary>The tx algorithm: margins from the self constraints, the primFontSz constraint as the starting size and the rule as the floor (autofit), anchor, bullets, alignment.</summary>
        private static void Text(AlgorithmAtom algorithm, DiagramShape shape, List<Constraint> constraints, List<RuleAtom> rules)
        {
            LayoutNodeAtom? node = algorithm.OwnerNode;
            var own = node?.Children.OfType<ConstraintAtom>().Select(c => c.Constraint).ToList() ?? [];
            shape.Margins = own.Where(c => c.Type is "lMarg" or "rMarg" or "tMarg" or "bMarg" && c.ForName.Length == 0 && c.RefType is "primFontSz" or "w" or "h")
                               .Select(c => (c.Type + ":" + c.RefType, c.Factor)).ToArray();
            double? font = null;
            foreach (Constraint constraint in constraints)
                if (constraint.Type == "primFontSz" && (constraint.ForName == shape.Name || constraint.ForName.Length == 0 && constraint.For == "none") && constraint.Value > 0)
                    font = constraint.Value;
            foreach (Constraint constraint in own)
                if (constraint.Type == "primFontSz" && constraint.ForName.Length == 0 && constraint.Value > 0) font = constraint.Value;
            shape.FontSize = Math.Min(Math.Max(font ?? 65, 1), 400);
            RuleAtom? floor = node?.Children.OfType<RuleAtom>().FirstOrDefault(r => r.Type == "primFontSz") ?? rules.FirstOrDefault(r => r.ForName == shape.Name && r.Type == "primFontSz");
            shape.MinimumFontSize = floor is { Value: > 0 } ? Math.Min(floor.Value, shape.FontSize) : shape.FontSize;
            shape.Autofit = true;
            shape.Anchor = algorithm.Parameter("txAnchorVert", "mid") switch { "t" => "t", "b" => "b", _ => "ctr" };
            int start = (int)algorithm.Number("stBulletLvl", 2) - 1;
            shape.BulletStartLevel = start;
            if (shape.Paragraphs.Count == 0) return;
            int baseLevel = shape.Paragraphs.Min(p => p.Level);
            shape.BulletList = shape.Paragraphs.Any(p => p.Level - baseLevel >= start);
            shape.CentreText = !shape.BulletList && !algorithm.Parameters.ContainsKey("parTxLTRAlign");
        }
    }

    // ── text fitting: the autofit LibreOffice turns on for diagram text, one size for each layout node's shapes ─────────

    private static class TextFitting
    {
        public static void FitAll(DiagramShape frame)
        {
            var shapes = new List<DiagramShape>();
            Gather(frame, shapes);
            foreach (IGrouping<string, DiagramShape> group in shapes.Where(s => s.HasText && s.Autofit).GroupBy(s => s.Name))
            {
                double size = group.Min(Fit);
                foreach (DiagramShape shape in group) shape.FontSize = size;
            }
        }

        private static void Gather(DiagramShape shape, List<DiagramShape> into)
        {
            into.Add(shape);
            foreach (DiagramShape child in shape.Children) Gather(child, into);
        }

        /// <summary>The largest size from the starting one down to the floor (whole points) at which the text fits the shape inside its margins.</summary>
        private static double Fit(DiagramShape shape)
        {
            for (double size = Math.Floor(shape.FontSize); size > shape.MinimumFontSize; size -= 1)
                if (Fits(shape, size)) return size;
            return Math.Min(shape.FontSize, shape.MinimumFontSize);
        }

        private static bool Fits(DiagramShape shape, double size)
        {
            (double left, double top, double right, double bottom) = Insets(shape, size);
            double width = shape.Width / EmuPerPoint - left - right, height = shape.Height / EmuPerPoint - top - bottom;
            if (width <= 0 || height <= 0) return false;
            int baseLevel = shape.Paragraphs.Min(p => p.Level);
            double used = 0;
            foreach ((XElement paragraph, int level) in shape.Paragraphs)
            {
                double paragraphSize = level > baseLevel && shape.BulletList ? size * 0.78 : size;
                double indent = shape.BulletList && level - baseLevel >= shape.BulletStartLevel ? 22.5 * (level - baseLevel - shape.BulletStartLevel + 1) : 0;
                double lineWidth = width - indent;
                string text = string.Concat(paragraph.Elements("r").Select(r => (string?)r.Element("t") ?? ""));
                int lines = 1;
                double current = 0;
                foreach (string word in text.Split(Space, StringSplitOptions.RemoveEmptyEntries))
                {
                    double wordWidth = Estimate(word, paragraphSize);
                    if (wordWidth > lineWidth) return false;
                    double withSpace = current == 0 ? wordWidth : current + paragraphSize * 0.25 + wordWidth;
                    if (withSpace > lineWidth) { lines++; current = wordWidth; }
                    else current = withSpace;
                }
                used += lines * paragraphSize * 1.2 * 0.9 + (used > 0 ? paragraphSize * 0.35 : 0);
            }
            return used <= height;
        }

        /// <summary>An average-width estimate of a word's width in points (no font metrics in the core: half an em per character, wider capitals).</summary>
        private static double Estimate(string word, double size)
        {
            double ems = 0;
            foreach (char c in word) ems += char.IsUpper(c) ? 0.62 : char.IsDigit(c) ? 0.51 : c is 'i' or 'l' or 'j' or 't' or 'f' or 'r' ? 0.3 : c is 'm' or 'w' ? 0.8 : 0.5;
            return ems * size;
        }
    }

    internal static (double Left, double Top, double Right, double Bottom) Insets(DiagramShape shape, double size)
    {
        double left = 0, top = 0, right = 0, bottom = 0;
        foreach ((string key, double factor) in shape.Margins)
        {
            string[] parts = key.Split(':');
            double value = parts[1] switch
            {
                "primFontSz" => factor * size,
                "w" => factor * shape.Width / EmuPerPoint,
                _ => factor * shape.Height / EmuPerPoint,
            };
            switch (parts[0]) { case "lMarg": left = value; break; case "tMarg": top = value; break; case "rMarg": right = value; break; default: bottom = value; break; }
        }
        return (left, top, right, bottom);
    }

    /// <summary>sortChildrenByZOrder (diagram.cxx): zOrderOff moves a shape that many places up.</summary>
    private static void SortByZOrder(DiagramShape shape)
    {
        List<DiagramShape> children = shape.Children;
        for (int i = 0; i < children.Count; i++) children[i].ZOrder = i;
        for (int i = 0; i < children.Count; i++)
        {
            int offset = children[i].ZOrderOffset;
            if (offset <= 0) continue;
            children[i].ZOrder += offset;
            for (int j = 0; j < offset && i + j + 1 < children.Count; j++) children[i + j + 1].ZOrder--;
        }
        var sorted = children.OrderBy(c => c.ZOrder).ToList();
        children.Clear();
        children.AddRange(sorted);
        foreach (DiagramShape child in children) SortByZOrder(child);
    }

    // ── output: dsp:sp elements, as Word writes the drawing part ──────────────────────────────────────────────────────────

    private static void Emit(DiagramShape shape, double offsetX, double offsetY, XElement tree, Diagram diagram)
    {
        double x = offsetX + shape.X, y = offsetY + shape.Y;
        if (!shape.IsGroup && (shape.Preset is not null || shape.HasText) && shape.Width > 0 && shape.Height > 0) tree.Add(Sp(shape, x, y, diagram));
        foreach (DiagramShape child in shape.Children) Emit(child, x, y, tree, diagram);
    }

    private static XElement Sp(DiagramShape shape, double x, double y, Diagram diagram)
    {
        var transform = new XElement("xfrm", new XElement("off", Attr("x", x), Attr("y", y)), new XElement("ext", Attr("cx", shape.Width), Attr("cy", shape.Height)));
        if (shape.Rotation != 0) transform.SetAttributeValue("rot", ((long)Math.Round(shape.Rotation * 60000)).ToString(CultureInfo.InvariantCulture));
        var properties = new XElement("spPr", transform);
        bool hidden = shape.HideGeometry || shape.Preset is null or "conn";
        properties.Add(new XElement("prstGeom", new XAttribute("prst", hidden ? "rect" : shape.Preset!), Adjustments(shape.Adjustments)));
        XElement? ownFill = shape.OwnProperties.LastOrDefault(e => e.Name.LocalName is "noFill" or "solidFill" or "gradFill" or "blipFill" or "pattFill");
        XElement? ownLine = shape.OwnProperties.LastOrDefault(e => e.Name.LocalName == "ln");
        if (hidden)
        {
            properties.Add(new XElement("noFill"));
            properties.Add(new XElement("ln", new XElement("noFill")));
        }
        else
        {
            if (ownFill is not null) properties.Add(new XElement(ownFill));
            if (ownLine is not null) properties.Add(new XElement(ownLine));
        }
        var sp = new XElement("sp", new XAttribute("modelId", shape.Presentation?.Id ?? ""), properties, Style(shape, diagram));
        if (shape.HasText) sp.Add(TextBody(shape));
        return sp;
    }

    private static IEnumerable<XElement> Adjustments(XElement? adjustments)
    {
        var list = new XElement("avLst");
        foreach (XElement adjustment in adjustments?.Elements("adj") ?? [])
        {
            int index = Int(adjustment.Attribute("idx"), 1);
            string value = ((long)Math.Round(Double(adjustment.Attribute("val"), 0) * 100000)).ToString(CultureInfo.InvariantCulture);
            if (index == 1) list.Add(new XElement("gd", new XAttribute("name", "adj"), new XAttribute("fmla", "val " + value)));
            list.Add(new XElement("gd", new XAttribute("name", "adj" + index.ToString(CultureInfo.InvariantCulture)), new XAttribute("fmla", "val " + value)));
        }
        yield return list;
    }

    /// <summary>The quickStyle's references for the style label, each with the colors part's colour for it (the n-th of the list, by its method).</summary>
    private static XElement Style(DiagramShape shape, Diagram diagram)
    {
        XElement? style = diagram.QuickStyle?.Elements("styleLbl").FirstOrDefault(s => (string?)s.Attribute("name") == shape.StyleLabel)?.Element("style");
        XElement? colors = diagram.Colors?.Elements("styleLbl").FirstOrDefault(s => (string?)s.Attribute("name") == shape.StyleLabel);
        int count = shape.Presentation?.PresStyleCount ?? 0;
        XElement Reference(string name, string fallbackIndex, string colorList)
        {
            XElement reference = style?.Element(name) is { } found ? new XElement(found) : new XElement(name, new XAttribute("idx", fallbackIndex));
            if (Color(colors?.Element(colorList), shape.ColorIndex, count) is { } color)
            {
                reference.RemoveNodes();
                reference.Add(color);
            }
            return reference;
        }
        return new XElement("style",
                            Reference("lnRef", "0", "linClrLst"), Reference("fillRef", "0", "fillClrLst"),
                            Reference("effectRef", "0", "effectClrLst"), Reference("fontRef", "minor", "txFillClrLst"));
    }

    /// <summary>The colour of a colour list for the index-th point of a label: repeat / cycle take index modulo the list, span spreads the list over the points.</summary>
    private static XElement? Color(XElement? list, int index, int count)
    {
        if (list is null) return null;
        var colors = list.Elements().ToList();
        if (colors.Count == 0) return null;
        if (index < 0) return new XElement(colors[colors.Count - 1]);
        string method = (string?)list.Attribute("meth") ?? "span";
        if (method == "span" && count > 1)
            return new XElement(colors[(int)Math.Round((double)Math.Min(index, count - 1) * (colors.Count - 1) / (count - 1))]);
        if (method == "span") return new XElement(colors[0]);
        return new XElement(colors[index % colors.Count]);
    }

    private static XElement TextBody(DiagramShape shape)
    {
        (double left, double top, double right, double bottom) = Insets(shape, shape.FontSize);
        var bodyProperties = new XElement("bodyPr", new XAttribute("wrap", "square"),
                                          Attr("lIns", left * EmuPerPoint), Attr("tIns", top * EmuPerPoint), Attr("rIns", right * EmuPerPoint), Attr("bIns", bottom * EmuPerPoint),
                                          new XAttribute("anchor", shape.Anchor));
        var body = new XElement("txBody", bodyProperties, new XElement("lstStyle"));
        int baseLevel = shape.Paragraphs.Min(p => p.Level);
        foreach ((XElement source, int level) in shape.Paragraphs)
        {
            int relative = level - baseLevel;
            bool bullet = shape.BulletList && relative >= shape.BulletStartLevel;
            double size = relative > 0 && shape.BulletList ? Math.Round(shape.FontSize * 0.78) : shape.FontSize;
            var paragraph = new XElement(source);
            XElement paragraphProperties = paragraph.Element("pPr") ?? new XElement("pPr");
            paragraphProperties.Remove();
            paragraphProperties.SetAttributeValue("lvl", relative.ToString(CultureInfo.InvariantCulture));
            if (paragraphProperties.Attribute("algn") is null) paragraphProperties.SetAttributeValue("algn", shape.CentreText ? "ctr" : "l");
            if (bullet)
            {
                long margin = 285750L * (relative - shape.BulletStartLevel + 1);
                paragraphProperties.SetAttributeValue("marL", margin.ToString(CultureInfo.InvariantCulture));
                paragraphProperties.SetAttributeValue("indent", "-285750");
            }
            paragraphProperties.AddFirst(new XElement("lnSpc", new XElement("spcPct", new XAttribute("val", "90000"))));
            paragraphProperties.Add(new XElement("spcAft", new XElement("spcPct", new XAttribute("val", "35000"))));
            paragraphProperties.Add(bullet ? new XElement("buChar", new XAttribute("char", "•")) : new XElement("buNone"));
            paragraph.AddFirst(paragraphProperties);
            string points = ((long)Math.Round(size * 100)).ToString(CultureInfo.InvariantCulture);
            foreach (XElement run in paragraph.Elements().Where(e => e.Name.LocalName is "r" or "fld" or "endParaRPr"))
            {
                XElement runProperties = run.Name.LocalName == "endParaRPr" ? run : run.Element("rPr") ?? AddFirst(run, new XElement("rPr"));
                if (runProperties.Attribute("sz") is null) runProperties.SetAttributeValue("sz", points);
            }
            body.Add(paragraph);
        }
        return body;
    }

    private static XElement AddFirst(XElement parent, XElement child)
    {
        parent.AddFirst(child);
        return child;
    }

    private static XAttribute Attr(string name, double value) => new(name, ((long)Math.Round(value)).ToString(CultureInfo.InvariantCulture));

    private static int Int(XAttribute? attribute, int fallback) => attribute is not null && int.TryParse(attribute.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int value) ? value : fallback;
    private static int Int(string text, int fallback) => int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out int value) ? value : fallback;
    private static double Double(XAttribute? attribute, double fallback)
        => attribute is not null && double.TryParse(attribute.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out double value) && !double.IsNaN(value) && !double.IsInfinity(value) ? value : fallback;
    private static bool IsTrue(XAttribute? attribute) => attribute?.Value is "1" or "true";
}
