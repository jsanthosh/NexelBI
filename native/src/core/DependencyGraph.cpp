#include "DependencyGraph.h"
#include <queue>

std::string DependencyGraph::key(const CellAddress& addr) {
    return addr.toString().toStdString();
}

void DependencyGraph::addDependency(const CellAddress& dependent, const CellAddress& dependency) {
    std::string depKey = key(dependent);
    std::string depOnKey = key(dependency);

    m_dependencies[depKey].insert(depOnKey);
    m_dependents[depOnKey].insert(depKey);
}

void DependencyGraph::removeDependencies(const CellAddress& cell) {
    std::string cellKey = key(cell);

    // Remove this cell from all its dependencies' dependent lists
    auto it = m_dependencies.find(cellKey);
    if (it != m_dependencies.end()) {
        for (const auto& depOn : it->second) {
            auto dit = m_dependents.find(depOn);
            if (dit != m_dependents.end()) {
                dit->second.erase(cellKey);
            }
        }
        m_dependencies.erase(it);
    }
}

std::vector<CellAddress> DependencyGraph::getDependents(const CellAddress& cell) const {
    std::vector<CellAddress> result;
    auto it = m_dependents.find(key(cell));
    if (it != m_dependents.end()) {
        for (const auto& k : it->second) {
            result.push_back(CellAddress::fromString(QString::fromStdString(k)));
        }
    }
    return result;
}

// BFS to find all cells that need recalculation, in topological order
std::vector<CellAddress> DependencyGraph::getRecalcOrder(const CellAddress& changed) const {
    std::vector<CellAddress> order;
    std::unordered_set<std::string> visited;
    std::queue<std::string> queue;

    std::string startKey = key(changed);
    auto it = m_dependents.find(startKey);
    if (it == m_dependents.end()) return order;

    for (const auto& dep : it->second) {
        if (visited.find(dep) == visited.end()) {
            queue.push(dep);
            visited.insert(dep);
        }
    }

    while (!queue.empty()) {
        std::string current = queue.front();
        queue.pop();
        order.push_back(CellAddress::fromString(QString::fromStdString(current)));

        auto dit = m_dependents.find(current);
        if (dit != m_dependents.end()) {
            for (const auto& dep : dit->second) {
                if (visited.find(dep) == visited.end()) {
                    queue.push(dep);
                    visited.insert(dep);
                }
            }
        }
    }

    return order;
}

bool DependencyGraph::hasCircularDependency(const CellAddress& cell) const {
    std::unordered_set<std::string> visited;
    std::string cellKey = key(cell);
    return detectCycle(cellKey, cellKey, visited);
}

bool DependencyGraph::detectCycle(const std::string& start, const std::string& current,
                                   std::unordered_set<std::string>& visited) const {
    auto it = m_dependencies.find(current);
    if (it == m_dependencies.end()) return false;

    for (const auto& dep : it->second) {
        if (dep == start) return true;
        if (visited.find(dep) == visited.end()) {
            visited.insert(dep);
            if (detectCycle(start, dep, visited)) return true;
        }
    }
    return false;
}

void DependencyGraph::clear() {
    m_dependents.clear();
    m_dependencies.clear();
}
