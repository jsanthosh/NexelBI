#ifndef DEPENDENCYGRAPH_H
#define DEPENDENCYGRAPH_H

#include <QString>
#include <string>
#include <unordered_map>
#include <unordered_set>
#include <vector>
#include "CellRange.h"

class DependencyGraph {
public:
    DependencyGraph() = default;

    void addDependency(const CellAddress& dependent, const CellAddress& dependency);
    void removeDependencies(const CellAddress& cell);
    std::vector<CellAddress> getDependents(const CellAddress& cell) const;
    std::vector<CellAddress> getRecalcOrder(const CellAddress& changed) const;
    bool hasCircularDependency(const CellAddress& cell) const;
    void clear();

private:
    static std::string key(const CellAddress& addr);

    // cell -> set of cells that depend on it
    std::unordered_map<std::string, std::unordered_set<std::string>> m_dependents;
    // cell -> set of cells it depends on
    std::unordered_map<std::string, std::unordered_set<std::string>> m_dependencies;

    bool detectCycle(const std::string& start, const std::string& current,
                     std::unordered_set<std::string>& visited) const;
};

#endif // DEPENDENCYGRAPH_H
